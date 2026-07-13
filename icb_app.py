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
NCR_CORRELATION_TARGET = 0.75
NCR_CORRELATION_TOLERANCE = 0.05  # valide si abs(correlation) dans [0.70, 0.80]

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
        self.project_name = row.get("Project Name") or ""
        self.cc_project_number = self._normalize_key(row.get("CC Project Number"))
        self.task_number = row.get("Task Number")
        self.employee_number = row.get("Employee Number")
        self.full_name = row.get("Full Name")
        self.emp_grade = row.get("Emp Grade")
        self.expenditure_type = row.get("Expenditure Type")
        self.uom = row.get("UOM")
        self.quantity = row.get("Quantity")
        br = row.get("Bill Rate")
        self.bill_rate = br * 8.0 if br is not None else None
        self.raw_cost = row.get("Raw Cost")
        self.bill_amount = row.get("Bill Amount")
        self.invoice_type = row.get("Invoice Type")
        self.date = excel_serial_to_date(row.get("Expenditure Item Date"))
        # Colonne de correlation pour la detection d'anomalie
        raw_corr = row.get("FULL NCR to Direct NCR Correlation")
        self.ncr_correlation = float(raw_corr) if raw_corr is not None else None

    @property
    def month_key(self):
        return self.date.strftime("%Y-%m") if self.date else "Inconnu"

    @property
    def has_anomaly(self):
        """True si abs(ncr_correlation) est en dehors de [0.70, 0.80]."""
        if self.ncr_correlation is None:
            return False
        val = abs(self.ncr_correlation)
        return val < (NCR_CORRELATION_TARGET - NCR_CORRELATION_TOLERANCE) or \
               val > (NCR_CORRELATION_TARGET + NCR_CORRELATION_TOLERANCE)

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

    def _get_items(self, month_filter=None, start_date=None, end_date=None):
        items = self.filtered_items
        if month_filter is not None:
            items = [i for i in items if i.month_key == month_filter]
        if start_date and end_date:
            try:
                start = datetime.strptime(start_date, "%Y-%m-%d").date()
                end = datetime.strptime(end_date, "%Y-%m-%d").date()
                items = [i for i in items if i.date and start <= i.date <= end]
            except ValueError:
                pass
        return items

    def hierarchy_data(self, month_filter=None, start_date=None, end_date=None):
        """Retourne la hierarchie CC Project -> Consultants en dict serialisable."""
        items = self._get_items(month_filter, start_date, end_date)
        hierarchy = OrderedDict()

        for item in items:
            cc_proj = item.cc_project_number
            emp = item.employee_number

            if cc_proj not in hierarchy:
                hierarchy[cc_proj] = {
                    "project_name": item.project_name,
                    "consultants": OrderedDict(),
                    "totals": {"working_days": 0.0, "raw_cost": 0.0, "bill_amount": 0.0},
                }

            cc_data = hierarchy[cc_proj]

            if emp not in cc_data["consultants"]:
                cc_data["consultants"][emp] = {
                    "name": item.full_name,
                    "grade": item.emp_grade,
                    "ncr": item.bill_rate or 0.0,
                    "working_days": 0.0,
                    "total_bill_amount": 0.0,
                    "total_raw_cost": 0.0,
                    "has_anomaly": False,
                }

            cons = cc_data["consultants"][emp]
            cons["ncr"] = item.bill_rate or cons["ncr"]
            cons["working_days"] += (item.quantity or 0) / 8.0
            cons["total_bill_amount"] += item.bill_amount or 0
            cons["total_raw_cost"] += item.raw_cost or 0
            # Un consultant est en anomalie si au moins une de ses lignes l'est
            if item.has_anomaly:
                cons["has_anomaly"] = True

            cc_data["totals"]["working_days"] += (item.quantity or 0) / 8.0
            cc_data["totals"]["raw_cost"] += item.raw_cost or 0
            cc_data["totals"]["bill_amount"] += item.bill_amount or 0

        # tri et serialisation
        result = []
        for cc_proj in sorted(hierarchy.keys()):
            cc = hierarchy[cc_proj]
            consultants_list = []
            for emp_id, cons in cc["consultants"].items():
                consultants_list.append({
                    "employee_id": emp_id,
                    "name": cons["name"],
                    "grade": cons["grade"],
                    "ncr": round(cons["ncr"], 2),
                    "working_days": round(cons["working_days"], 2),
                    "total_bill_amount": round(cons["total_bill_amount"], 2),
                    "total_raw_cost": round(cons["total_raw_cost"], 2),
                    "has_anomaly": cons["has_anomaly"],
                })
            result.append({
                "cc_project_number": cc_proj,
                "project_name": cc["project_name"],
                "consultants": consultants_list,
                "totals": {k: round(v, 2) for k, v in cc["totals"].items()},
            })
        return result

    def consultant_transactions(self, employee_number, cc_project, month_filter=None):
        items = self._get_items(month_filter)
        txns = [
            i for i in items
            if i.employee_number == employee_number and i.cc_project_number == cc_project
        ]
        result = []
        for t in txns:
            result.append({
                "item_id": t.item_id,
                "date": t.date.strftime("%Y-%m-%d") if t.date else "",
                "expenditure_type": t.expenditure_type,
                "invoice_type": t.invoice_type or "",
                "working_days": round((t.quantity or 0) / 8.0, 2),
                "ncr": round(t.bill_rate, 2) if t.bill_rate else 0,
                "raw_cost": round(t.raw_cost, 2) if t.raw_cost else 0,
                "bill_amount": round(t.bill_amount, 2) if t.bill_amount else 0,
                "ncr_correlation": round(t.ncr_correlation, 4) if t.ncr_correlation is not None else None,
                "has_anomaly": t.has_anomaly,
            })
        return result

    def export_to_excel(self, start_date, end_date):
        import io
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment

        # 1. Obtenir les items dans l'intervalle pour determiner les mois concernes
        items_in_range = self._get_items(start_date=start_date, end_date=end_date)
        months_in_range = sorted(list(set(i.month_key for i in items_in_range)))

        wb = Workbook()
        # Supprimer la feuille par defaut si on va creer des feuilles specifiques
        default_ws = wb.active

        # Styles
        header_fill = PatternFill(start_color="1A3A5C", end_color="1A3A5C", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        cc_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        cc_font = Font(bold=True)
        right_align = Alignment(horizontal="right")

        if not months_in_range:
            default_ws.title = "Aucune donnee"
            default_ws.append(["Aucune donnee dans cet intervalle"])
        else:
            for idx, month in enumerate(months_in_range):
                if idx == 0:
                    ws = default_ws
                    ws.title = month
                else:
                    ws = wb.create_sheet(title=month)

                data = self.hierarchy_data(month_filter=month, start_date=start_date, end_date=end_date)

                # Headers : Projet Name, CC Project Number, et les totaux
                headers = ["Project Name", "CC Project Number", "Working Days", "Bill Amount", "Raw Cost"]
                ws.append(headers)
                for col in range(1, 6):
                    cell = ws.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    if col >= 3:
                        cell.alignment = right_align

                current_row = 2
                for cc in data:
                    ws.append([
                        cc['project_name'] or "",
                        cc['cc_project_number'],
                        cc['totals']['working_days'],
                        cc['totals']['bill_amount'],
                        cc['totals']['raw_cost']
                    ])
                    for col in range(1, 6):
                        cell = ws.cell(row=current_row, column=col)
                        cell.fill = cc_fill
                        cell.font = cc_font
                        if col >= 3:
                            cell.alignment = right_align
                    current_row += 1

                # Adjust columns width
                ws.column_dimensions['A'].width = 40
                ws.column_dimensions['B'].width = 25
                ws.column_dimensions['C'].width = 15
                ws.column_dimensions['D'].width = 15
                ws.column_dimensions['E'].width = 15

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer


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


@app.route("/api/export")
def export_excel():
    global pipeline
    import os
    if pipeline is None:
        return jsonify({"error": "Aucun fichier charge."}), 400
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")
    if not start_date or not end_date:
        return jsonify({"error": "Dates manquantes."}), 400
    
    try:
        buffer = pipeline.export_to_excel(start_date, end_date)
        
        # Build the desktop path
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        filename = f"icb_{start_date}_{end_date}.xlsx"
        filepath = os.path.join(desktop, filename)
        
        with open(filepath, "wb") as f:
            f.write(buffer.getvalue())
            
        return jsonify({
            "success": True,
            "message": "Exportation avec succès",
            "path": filepath
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def start_server(port=5050):
    """Lance le serveur Flask (utilise par le wrapper desktop)."""
    import logging
    log = logging.getLogger("werkzeug")
    log.setLevel(logging.ERROR)
    app.run(debug=False, port=port, use_reloader=False)