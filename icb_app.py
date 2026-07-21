"""
Application ICB - Capgemini
Backend Flask : charge le fichier Excel, filtre, groupe, et expose les donnees
via des endpoints JSON pour l'interface web.
"""

import os
import io
import math
import tempfile
import pandas as pd
from datetime import datetime, date, timedelta
from collections import OrderedDict
from flask import Flask, request, jsonify, render_template
from openpyxl import load_workbook

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
        self.forecast_data = {}  # (cc_proj, emp, month_key) -> {"ncr": float, "working_days": float, "bill_amount": float}

    def load(self, path):
        df_raw = fast_load_excel(path)
        self.items = [ExpenditureItem(row) for row in df_raw.to_dict("records")]

    def load_forecast(self, path):
        """Parse un fichier Forecast et remplit self.forecast_data.
        
        Structure du fichier Forecast :
        - Row 1 : Noms des mois aux colonnes 7,13,19,...,73 (1-indexed)
        - Row 6 : En-têtes de colonnes (Grade, F.RC, Vaca, Prod, Cost, FTE) x 12 mois
        - Row 7+ : Données consultants
          - Col 1 (0-idx 0) : Project Num  (seulement sur la 1ère ligne du groupe)
          - Col 2 (0-idx 1) : Project Name
          - Col 3 (0-idx 2) : GGID
          - Col 4 (0-idx 3) : Employee Name
          - Chaque mois : 6 colonnes à partir de 0-idx 6, pas de 6
            Grade(+0), F.RC(+1), Vaca(+2), Prod(+3), Cost(+4), FTE(+5)
        """
        import openpyxl
        from datetime import datetime

        print("\n" + "=" * 70)
        print("[FORECAST IMPORT] Debut du chargement du fichier Forecast")
        print("=" * 70)

        # =================================================================
        # 1. VALIDATION STRUCTURELLE DE BASE
        #    Ouverture en mode normal (pas read_only) pour acceder
        #    aux merged_cells et aux dimensions reelles.
        # =================================================================
        print("\n--- 1. VALIDATION STRUCTURELLE DE BASE ---")

        wb_check = openpyxl.load_workbook(path, data_only=True)
        ws_check = wb_check.active

        print(f"  Feuille active : '{ws_check.title}'")
        print(f"  Dimensions declarees : {ws_check.dimensions}")

        # Verification des en-tetes ligne 6
        expected_headers = {
            1: "Project Num",
            2: "Project Name",
            3: "GGID",
            4: "Employee Name",
            7: "Grade",
        }
        headers_ok = True
        for col_1idx, expected_name in expected_headers.items():
            actual = ws_check.cell(row=6, column=col_1idx).value
            if actual != expected_name:
                print(f"  [WARNING] Colonne {col_1idx} (ligne 6) : attendu "
                      f"'{expected_name}', trouve '{actual}'")
                headers_ok = False
            else:
                print(f"  [OK] Colonne {col_1idx} : '{actual}'")
        if headers_ok:
            print("  => Tous les en-tetes valides.")

        total_cols = ws_check.max_column
        expected_cols = 87
        print(f"  Nombre de colonnes detectees : {total_cols} (attendu : {expected_cols})")
        if total_cols != expected_cols:
            print(f"  [WARNING] Nombre de colonnes inattendu ({total_cols} vs {expected_cols})")

        # =================================================================
        # 2. DETECTION DES BLOCS PROJETS (via plages fusionnees colonne A)
        # =================================================================
        print("\n--- 2. DETECTION DES BLOCS PROJETS ---")

        merged_ranges_a = sorted(
            [mr for mr in ws_check.merged_cells.ranges if mr.min_col == 1],
            key=lambda r: r.min_row,
        )

        detected_blocks = []   # [(proj_num, start_row, end_row, n_cons, has_total)]

        if merged_ranges_a:
            print(f"  {len(merged_ranges_a)} plage(s) fusionnee(s) en colonne A :")
            for mr in merged_ranges_a:
                anchor_val = ws_check.cell(row=mr.min_row, column=1).value
                n_cons = mr.max_row - mr.min_row + 1
                # Verifier la ligne Total juste apres le bloc
                total_row = mr.max_row + 1
                total_ggid = ws_check.cell(row=total_row, column=3).value
                total_emp  = ws_check.cell(row=total_row, column=4).value
                has_total = (total_ggid is None and total_emp is None)
                detected_blocks.append(
                    (anchor_val, mr.min_row, mr.max_row, n_cons, has_total)
                )
                print(
                    f"    Projet '{anchor_val}' : lignes {mr.min_row}-{mr.max_row} "
                    f"({n_cons} consultants), "
                    f"ligne Total={total_row} "
                    f"{'[detectee et exclue]' if has_total else '[NON detectee — WARNING]'}"
                )
        else:
            print("  [WARNING] Aucune plage fusionnee trouvee en colonne A.")
            print("  Tentative de detection par rupture de valeur :")
            proj_rows = {}
            for row_cells in ws_check.iter_rows(min_row=7, min_col=1, max_col=4,
                                                 values_only=False):
                val = row_cells[0].value
                if val is not None:
                    proj_rows.setdefault(val, []).append(row_cells[0].row)
            if proj_rows:
                for proj_val, rows in proj_rows.items():
                    print(f"    Projet '{proj_val}' : present aux lignes {rows}")
            else:
                print("    [WARNING] Aucun Project Num trouve du tout.")

        wb_check.close()

        # =================================================================
        # Parsing principal (read_only pour la performance)
        # =================================================================
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.active
            current_year = datetime.now().year
            current_proj = None

            # Compteurs de diagnostic
            projects_found = set()          # projets normalises
            consultants_by_project = {}     # proj_norm -> set((ggid_brut, ggid_norm))
            rows_processed = 0
            rows_skipped_total = 0

            for row in ws.iter_rows(min_row=7, values_only=True):
                proj_num = row[0]   # Col A (0-idx 0)
                ggid = row[2]       # Col C (0-idx 2)

                # Mettre à jour le projet courant si la ligne en contient un
                if proj_num is not None:
                    current_proj = proj_num

                # Ignorer les lignes de total (GGID vide = ligne résumé projet)
                if ggid is None:
                    rows_skipped_total += 1
                    continue

                rows_processed += 1
                norm_proj = ExpenditureItem._normalize_key(current_proj)
                norm_ggid = str(ggid).strip().replace("6427_", "")

                projects_found.add(norm_proj)
                consultants_by_project.setdefault(norm_proj, set()).add(
                    (str(ggid).strip(), norm_ggid)
                )

                for m in range(12):
                    month_str = f"{m+1:02d}"
                    month_key = f"{current_year}-{month_str}"

                    # Chaque bloc mois = 6 colonnes, démarrant à 0-idx 6
                    # Mois 0 (Jan) : cols 6..11, Mois 1 (Feb) : cols 12..17, etc.
                    base_idx = 6 + (m * 6)

                    if base_idx + 4 >= len(row):
                        break

                    # Grade = base_idx + 0
                    # F.RC  = base_idx + 1  → NCR forecast
                    # Vaca  = base_idx + 2
                    # Prod  = base_idx + 3  → Working Days forecast
                    # Cost  = base_idx + 4  → Bill Amount forecast
                    # FTE   = base_idx + 5
                    val_frc = row[base_idx + 1]
                    val_prod = row[base_idx + 3]
                    val_cost = row[base_idx + 4]

                    def to_float(val):
                        try:
                            if val is None or str(val).strip() == "" or str(val).strip() == "-":
                                return 0.0
                            return float(val)
                        except ValueError:
                            return 0.0

                    frc = to_float(val_frc)
                    prod = to_float(val_prod)
                    cost = to_float(val_cost)

                    if frc != 0 or prod != 0 or cost != 0:
                        self.forecast_data[(norm_proj, norm_ggid, month_key)] = {
                            "ncr": frc,
                            "working_days": prod,
                            "bill_amount": cost
                        }

            # =============================================================
            # 3. MAPPING AVEC LE FICHIER ICB
            # =============================================================
            print("\n--- 3. MAPPING AVEC LE FICHIER ICB ---")

            # Extraire les projets et consultants connus de l'ICB
            icb_projects = set()
            icb_employees_by_project = {}   # proj_norm -> set(emp_norm)
            for item in self.filtered_items:
                p = ExpenditureItem._normalize_key(item.cc_project_number)
                e = str(item.employee_number).strip().replace("6427_", "")
                icb_projects.add(p)
                icb_employees_by_project.setdefault(p, set()).add(e)

            # -- Projets --
            matched_projects = projects_found & icb_projects
            unmatched_projects = projects_found - icb_projects

            print(f"  Projets Forecast : {len(projects_found)}")
            print(f"  Projets ICB      : {len(icb_projects)}")
            print(f"  Projets matches  : {len(matched_projects)}/{len(projects_found)}")
            if matched_projects:
                for p in sorted(matched_projects):
                    print(f"    [MATCH]    {p}")
            if unmatched_projects:
                print(f"  [WARNING] {len(unmatched_projects)} projet(s) Forecast "
                      f"NON trouves dans l'ICB :")
                for p in sorted(unmatched_projects):
                    print(f"    [NO MATCH] {p}")

            # -- Consultants --
            total_matched_cons = 0
            total_unmatched_cons = 0
            for proj_norm, cons_set in consultants_by_project.items():
                icb_emps = icb_employees_by_project.get(proj_norm, set())
                for raw_ggid, norm_ggid in sorted(cons_set):
                    if norm_ggid in icb_emps:
                        total_matched_cons += 1
                    else:
                        total_unmatched_cons += 1
                        print(
                            f"  [NO MATCH] Consultant brut='{raw_ggid}', "
                            f"normalise='{norm_ggid}' dans projet '{proj_norm}' "
                            f"— absent de l'ICB"
                        )

            total_cons = total_matched_cons + total_unmatched_cons
            print(f"  Consultants matches : {total_matched_cons}/{total_cons}")
            if total_unmatched_cons:
                print(f"  [WARNING] {total_unmatched_cons} consultant(s) non matches")

            # =============================================================
            # 4. RESULTAT FINAL
            # =============================================================
            cells_filled = len(self.forecast_data) * 3   # NCR + WD + BA par entree

            print("\n--- 4. RESULTAT FINAL ---")
            print(f"  Lignes consultants traitees : {rows_processed}")
            print(f"  Lignes Total/vides ignorees : {rows_skipped_total}")
            print(f"  Entrees forecast_data       : {len(self.forecast_data)}")
            print(f"  Cellules FORECAST remplies   : {cells_filled} "
                  f"(NCR + Working Days + Bill Amount)")
            print(
                f"\n[FORECAST IMPORT] {len(projects_found)} projets detectes, "
                f"{total_matched_cons}/{total_cons} consultants matches, "
                f"{cells_filled} cellules remplies"
            )
            print("=" * 70 + "\n")

        finally:
            wb.close()

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
                    "totals": {
                        "working_days": 0.0, "raw_cost": 0.0, "bill_amount": 0.0,
                        "forecast_working_days": 0.0, "forecast_bill_amount": 0.0, "forecast_ncr": 0.0
                    },
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
                    "forecast_working_days": 0.0,
                    "forecast_bill_amount": 0.0,
                    "forecast_ncr": 0.0,
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

        # Integration des donnees Forecast
        # Determiner les mois actifs pour le forecast
        active_months = []
        if month_filter:
            active_months = [month_filter]
        else:
            if start_date and end_date:
                start = datetime.strptime(start_date, "%Y-%m-%d").date()
                end = datetime.strptime(end_date, "%Y-%m-%d").date()
                for m in self.available_months:
                    m_date = datetime.strptime(m + "-01", "%Y-%m-%d").date()
                    if start <= m_date <= end: # Approximation simple
                        active_months.append(m)
            else:
                active_months = self.available_months

        for cc_proj, cc_data in hierarchy.items():
            for emp, cons in cc_data["consultants"].items():
                norm_proj = ExpenditureItem._normalize_key(cc_proj)
                norm_emp = str(emp).strip().replace("6427_", "")
                
                sum_frc = 0.0
                sum_prod = 0.0
                sum_cost = 0.0
                frc_count = 0
                for m in active_months:
                    key = (norm_proj, norm_emp, m)
                    if key in self.forecast_data:
                        fd = self.forecast_data[key]
                        sum_frc += fd["ncr"]
                        sum_prod += fd["working_days"]
                        sum_cost += fd["bill_amount"]
                        if fd["ncr"] != 0:
                            frc_count += 1
                
                cons["forecast_working_days"] = sum_prod
                cons["forecast_bill_amount"] = sum_cost
                # NCR Forecast = moyenne des F.RC non nulles, ou calcul dérivé
                if frc_count > 0:
                    cons["forecast_ncr"] = sum_frc / frc_count
                elif sum_prod > 0:
                    cons["forecast_ncr"] = (sum_cost / sum_prod) * 8.0
                else:
                    cons["forecast_ncr"] = 0.0

                cc_data["totals"]["forecast_working_days"] += cons["forecast_working_days"]
                cc_data["totals"]["forecast_bill_amount"] += cons["forecast_bill_amount"]
            
            # Calcul du NCR global pour le projet (informatif)
            if cc_data["totals"]["forecast_working_days"] > 0:
                cc_data["totals"]["forecast_ncr"] = (cc_data["totals"]["forecast_bill_amount"] / cc_data["totals"]["forecast_working_days"]) * 8.0
            else:
                cc_data["totals"]["forecast_ncr"] = 0.0

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
                    "forecast_ncr": round(cons["forecast_ncr"], 2),
                    "forecast_working_days": round(cons["forecast_working_days"], 2),
                    "forecast_bill_amount": round(cons["forecast_bill_amount"], 2),
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

    def export_to_excel(self, month_filter, start_date, end_date):
        import io
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment

        # 1. Obtenir les items dans l'intervalle pour determiner les mois concernes
        items_in_range = self._get_items(month_filter=month_filter, start_date=start_date, end_date=end_date)
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

                # Headers : Project Name, CC Project Number, Working Days, Bill Amount, Forecast, Gap, Raw Cost, Comment
                headers = ["Project Name", "CC Project Number", "Working Days", "Bill Amount", "Forecast bill amount", "Gap", "Raw Cost", "Comment"]
                ws.append(headers)
                for col in range(1, 9):
                    cell = ws.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    if col >= 3 and col <= 7:
                        cell.alignment = right_align

                current_row = 2
                for cc in data:
                    ws.append([
                        cc['project_name'] or "",
                        cc['cc_project_number'],
                        cc['totals']['working_days'],
                        cc['totals']['bill_amount'],
                        "", # Forecast
                        f"=D{current_row}-E{current_row}", # Gap
                        cc['totals']['raw_cost'],
                        ""  # Comment
                    ])
                    for col in range(1, 9):
                        cell = ws.cell(row=current_row, column=col)
                        cell.fill = cc_fill
                        cell.font = cc_font
                        if col >= 3 and col <= 7:
                            cell.alignment = right_align
                    current_row += 1

                # Adjust columns width
                ws.column_dimensions['A'].width = 40
                ws.column_dimensions['B'].width = 25
                ws.column_dimensions['C'].width = 15
                ws.column_dimensions['D'].width = 15
                ws.column_dimensions['E'].width = 20
                ws.column_dimensions['F'].width = 15
                ws.column_dimensions['G'].width = 15
                ws.column_dimensions['H'].width = 30

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

@app.route("/api/upload_forecast", methods=["POST"])
def upload_forecast():
    if "file" not in request.files:
        return jsonify({"error": "Aucun fichier envoye."}), 400
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "Le fichier doit etre au format .xlsx"}), 400

    # sauvegarde temporaire
    upload_dir = os.path.join(os.path.dirname(__file__), "uploads")
    os.makedirs(upload_dir, exist_ok=True)
    path = os.path.join(upload_dir, "current_forecast.xlsx")
    f.save(path)

    try:
        pipeline.load_forecast(path)
        return jsonify({"message": "Fichier Forecast charge avec succes."}), 200
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/hierarchy")
def get_hierarchy():
    global pipeline
    if pipeline is None:
        return jsonify({"error": "Aucun fichier charge."}), 400
    month = request.args.get("month", None)
    if month == "all" or month == "":
        month = None
    start_date = request.args.get("start_date", None)
    end_date = request.args.get("end_date", None)
    data = pipeline.hierarchy_data(month_filter=month, start_date=start_date, end_date=end_date)
    return jsonify(data)

@app.route("/api/months")
def get_months():
    global pipeline
    if pipeline is None:
        return jsonify({"error": "Aucun fichier charge."}), 400
    start_date = request.args.get("start_date", None)
    end_date = request.args.get("end_date", None)
    items = pipeline._get_items(start_date=start_date, end_date=end_date)
    months = sorted(list(set(i.month_key for i in items)))
    return jsonify(months)


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
    start_date = request.args.get("start_date", None)
    end_date = request.args.get("end_date", None)
    
    if not emp or not cc_proj:
        return jsonify({"error": "Parametres employee et cc_project requis."}), 400
        
    items = pipeline._get_items(month_filter=month, start_date=start_date, end_date=end_date)
    txns = [
        i for i in items
        if i.employee_number == emp and i.cc_project_number == cc_proj
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
    return jsonify(result)


@app.route("/api/export")
def export_excel():
    global pipeline
    import os
    if pipeline is None:
        return jsonify({"error": "Aucun fichier charge."}), 400
    
    month = request.args.get("month", None)
    if month == "all" or month == "":
        month = None
    start_date = request.args.get("start_date", None)
    end_date = request.args.get("end_date", None)
    
    try:
        buffer = pipeline.export_to_excel(month_filter=month, start_date=start_date, end_date=end_date)
        
        import os
        if start_date and end_date:
            filename = f"icb_{start_date}_{end_date}.xlsx"
        elif month:
            filename = f"icb_{month}.xlsx"
        else:
            filename = "icb_export.xlsx"
        filepath = None
        
        try:
            import webview
            if webview.windows:
                window = webview.windows[0]
                result = window.create_file_dialog(
                    webview.FileDialog.SAVE,
                    save_filename=filename,
                    file_types=('Excel Files (*.xlsx)',)
                )
                if result and isinstance(result, (tuple, list)) and len(result) > 0:
                    filepath = result[0]
                elif result and isinstance(result, str):
                    filepath = result
                else:
                    return jsonify({"error": "Exportation annulée."}), 400
        except ImportError:
            pass
            
        if not filepath:
            # Build the desktop path as fallback
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
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
    app.run(debug=True, port=port, use_reloader=False)