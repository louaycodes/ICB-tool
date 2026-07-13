"""
Application ICB - Capgemini
Backend Flask : charge le fichier Excel, filtre, groupe, et expose les donnees
via des endpoints JSON pour l'interface web.
"""

import os
import json
from datetime import datetime, timedelta
from collections import defaultdict, OrderedDict

import pandas as pd
from openpyxl import load_workbook
from flask import Flask, render_template, request, jsonify

# --- Parametres metier ---
EXPENDITURE_TYPES = [
    "RC_Time Std",
    "RC_Overtime Std Cost",
    "RC_Overtime Std Cost x2",
    "RC_Overtime Std Cost x1.25",
    "RC_Overtime Std Cost x1.5",
]
UOM_FILTER = "HOURS"
BILL_RATIO = 0.75
TOLERANCE_PCT = 0.02

# --- Flask app ---
app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 Mo max pour l'upload

# stockage en memoire du pipeline (une seule session)
pipeline = None


# ---------------------------------------------------------------------------
# Lecture rapide de gros fichiers Excel
# ---------------------------------------------------------------------------
def fast_load_excel(path, sheet_name=0):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if isinstance(sheet_name, str) else wb.worksheets[sheet_name]
    header_row = next(ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=200, values_only=True))
    ncols = max((i for i, v in enumerate(header_row) if v is not None), default=-1) + 1
    if ncols == 0:
        wb.close()
        raise ValueError("Feuille vide ou en-tete introuvable.")
    header = list(header_row[:ncols])
    rows = ws.iter_rows(min_row=2, min_col=1, max_col=ncols, values_only=True)
    data = [r for r in rows if any(v is not None for v in r)]
    wb.close()
    return pd.DataFrame(data, columns=header)


def excel_serial_to_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
        except (OverflowError, ValueError):
            return None
    return None


# ---------------------------------------------------------------------------
# Modele metier
# ---------------------------------------------------------------------------
class ExpenditureItem:
    """Une ligne de depense/facturation."""

    def __init__(self, row: dict):
        self.item_id = row.get("Expenditure Item Id")
        self.project_number = self._normalize_key(row.get("Project Number"))
        self.cc_project_number = self._normalize_key(row.get("CC Project Number"))
        self.task_number = row.get("Task Number")
        self.employee_number = row.get("Employee Number")
        self.full_name = row.get("Full Name")
        self.emp_grade = row.get("Emp Grade")
        self.expenditure_type = row.get("Expenditure Type")
        self.uom = row.get("UOM")
        self.quantity = row.get("Quantity")
        self.bill_rate = row.get("Bill Rate")
        self.raw_cost = row.get("Raw Cost")
        self.bill_amount = row.get("Bill Amount")
        self.invoice_type = row.get("Invoice Type")
        self.date = excel_serial_to_date(row.get("Expenditure Item Date"))

    @property
    def month_key(self):
        return self.date.strftime("%Y-%m") if self.date else "Inconnu"

    @property
    def expected_bill(self):
        if self.bill_rate is not None and self.quantity is not None:
            return self.bill_rate * self.quantity / 8
        return 0.0

    @staticmethod
    def _normalize_key(value):
        """Force toutes les cles projet/CC-projet en texte propre (evite le
        melange str/int/None qui fait planter sorted() plus loin)."""
        if value is None:
            return "N/A"
        if isinstance(value, float) and value.is_integer():
            return str(int(value))   # 123.0 -> "123", evite les doublons avec "123"
        return str(value).strip()

    def matches_filter(self):
        return self.expenditure_type in EXPENDITURE_TYPES and self.uom == UOM_FILTER


# ---------------------------------------------------------------------------
# Pipeline de traitement
# ---------------------------------------------------------------------------
class ICBPipeline:
    def __init__(self):
        self.items = []
        self.filtered_items = []
        self.available_months = []

    def load(self, path):
        df_raw = fast_load_excel(path)
        self.items = [ExpenditureItem(row) for row in df_raw.to_dict("records")]

    def process(self):
        self.filtered_items = [i for i in self.items if i.matches_filter()]
        months_set = set()
        for item in self.filtered_items:
            months_set.add(item.month_key)
        self.available_months = sorted(months_set)

    def _get_items(self, month_filter=None):
        if month_filter is None:
            return self.filtered_items
        return [i for i in self.filtered_items if i.month_key == month_filter]

    def hierarchy_data(self, month_filter=None):
        """Retourne la hierarchie Project -> CC Project -> Consultants en dict serialisable."""
        items = self._get_items(month_filter)
        hierarchy = OrderedDict()

        for item in items:
            proj = item.project_number
            cc_proj = item.cc_project_number
            emp = item.employee_number

            if proj not in hierarchy:
                hierarchy[proj] = OrderedDict()

            if cc_proj not in hierarchy[proj]:
                hierarchy[proj][cc_proj] = {
                    "consultants": OrderedDict(),
                    "totals": {"hours": 0.0, "raw_cost": 0.0, "bill_amount": 0.0, "expected_bill": 0.0},
                }

            cc_data = hierarchy[proj][cc_proj]

            if emp not in cc_data["consultants"]:
                cc_data["consultants"][emp] = {
                    "name": item.full_name,
                    "grade": item.emp_grade,
                    "ncr": item.bill_rate or 0.0,
                    "total_hours": 0.0,
                    "total_bill_amount": 0.0,
                    "total_raw_cost": 0.0,
                    "total_expected_bill": 0.0,
                }

            cons = cc_data["consultants"][emp]
            cons["ncr"] = item.bill_rate or cons["ncr"]
            cons["total_hours"] += item.quantity or 0
            cons["total_bill_amount"] += item.bill_amount or 0
            cons["total_raw_cost"] += item.raw_cost or 0
            cons["total_expected_bill"] += item.expected_bill

            cc_data["totals"]["hours"] += item.quantity or 0
            cc_data["totals"]["raw_cost"] += item.raw_cost or 0
            cc_data["totals"]["bill_amount"] += item.bill_amount or 0
            cc_data["totals"]["expected_bill"] += item.expected_bill

        # tri
        result = []
        for proj in sorted(hierarchy.keys()):
            proj_data = {
                "project_number": proj,
                "cc_projects": [],
                "totals": {"hours": 0.0, "raw_cost": 0.0, "bill_amount": 0.0, "expected_bill": 0.0},
            }
            for cc_proj in sorted(hierarchy[proj].keys()):
                cc = hierarchy[proj][cc_proj]
                consultants_list = []
                for emp_id, cons in cc["consultants"].items():
                    # verifie l'anomalie : raw_cost attendu = 0.75 * bill_amount
                    expected_raw = BILL_RATIO * abs(cons["total_bill_amount"])
                    actual_raw = abs(cons["total_raw_cost"])
                    has_anomaly = abs(actual_raw - expected_raw) > 2.0
                    consultants_list.append({
                        "employee_id": emp_id,
                        "name": cons["name"],
                        "grade": cons["grade"],
                        "ncr": round(cons["ncr"], 2),
                        "total_hours": round(cons["total_hours"], 2),
                        "total_bill_amount": round(cons["total_bill_amount"], 2),
                        "total_raw_cost": round(cons["total_raw_cost"], 2),
                        "total_expected_bill": round(cons["total_expected_bill"], 2),
                        "expected_raw_cost": round(expected_raw, 2),
                        "has_anomaly": has_anomaly,
                    })
                cc_entry = {
                    "cc_project_number": cc_proj,
                    "consultants": consultants_list,
                    "totals": {k: round(v, 2) for k, v in cc["totals"].items()},
                }
                proj_data["cc_projects"].append(cc_entry)
                for k in proj_data["totals"]:
                    proj_data["totals"][k] += cc["totals"][k]
            proj_data["totals"] = {k: round(v, 2) for k, v in proj_data["totals"].items()}
            result.append(proj_data)
        return result

    def consultant_transactions(self, employee_number, cc_project, month_filter=None):
        items = self._get_items(month_filter)
        txns = [
            i for i in items
            if i.employee_number == employee_number and i.cc_project_number == cc_project
        ]
        result = []
        for t in txns:
            # L'anomalie s'applique uniquement à RC_Time Std
            has_anomaly = False
            if t.expenditure_type == "RC_Time Std":
                expected_raw = BILL_RATIO * abs(t.bill_amount or 0)
                actual_raw = abs(t.raw_cost or 0)
                has_anomaly = abs(actual_raw - expected_raw) > 2.0
            
            result.append({
                "item_id": t.item_id,
                "date": t.date.strftime("%Y-%m-%d") if t.date else "",
                "expenditure_type": t.expenditure_type,
                "invoice_type": t.invoice_type or "",
                "quantity": round(t.quantity, 2) if t.quantity else 0,
                "ncr": round(t.bill_rate, 2) if t.bill_rate else 0,
                "raw_cost": round(t.raw_cost, 2) if t.raw_cost else 0,
                "bill_amount": round(t.bill_amount, 2) if t.bill_amount else 0,
                "expected_bill": round(t.expected_bill, 2),
                "has_anomaly": has_anomaly,
            })
        return result


# ---------------------------------------------------------------------------
# Routes Flask
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload_file():
    global pipeline
    if "file" not in request.files:
        return jsonify({"error": "Aucun fichier envoye."}), 400
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "Le fichier doit etre au format .xlsx"}), 400

    # sauvegarde temporaire
    upload_dir = os.path.join(os.path.dirname(__file__), "uploads")
    os.makedirs(upload_dir, exist_ok=True)
    path = os.path.join(upload_dir, "current_icb.xlsx")
    f.save(path)

    try:
        pipe = ICBPipeline()
        pipe.load(path)
        pipe.process()
        pipeline = pipe
    except Exception as exc:
        return jsonify({"error": f"Erreur lors du chargement : {str(exc)}"}), 500

    return jsonify({
        "total_lines": len(pipeline.items),
        "filtered_lines": len(pipeline.filtered_items),
        "available_months": pipeline.available_months,
    })


@app.route("/api/hierarchy")
def get_hierarchy():
    global pipeline
    if pipeline is None:
        return jsonify({"error": "Aucun fichier charge."}), 400
    month = request.args.get("month", None)
    if month == "all" or month == "":
        month = None
    data = pipeline.hierarchy_data(month)
    return jsonify(data)


@app.route("/api/transactions")
def get_transactions():
    global pipeline
    if pipeline is None:
        return jsonify({"error": "Aucun fichier charge."}), 400
    emp = request.args.get("employee")
    cc_proj = request.args.get("cc_project")
    month = request.args.get("month", None)
    if month == "all" or month == "":
        month = None
    if not emp or not cc_proj:
        return jsonify({"error": "Parametres employee et cc_project requis."}), 400
    data = pipeline.consultant_transactions(emp, cc_proj, month)
    return jsonify(data)


def start_server(port=5050):
    """Lance le serveur Flask (utilise par le wrapper desktop)."""
    import logging
    log = logging.getLogger("werkzeug")
    log.setLevel(logging.ERROR)
    app.run(debug=False, port=port, use_reloader=False)