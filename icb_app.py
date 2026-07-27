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

# Seuils de tolerance pour la detection d'ecarts ICB vs FORECAST
# (surlignage rouge des cellules + generation des commentaires export)
GAP_TOLERANCE_AMOUNT = 2    # en unites monetaires (EUR/USD) — pour Bill Amount et NCR
GAP_TOLERANCE_DAYS   = 0.5  # en jours — pour Working Days

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
        self.forecast_names = {}  # (norm_proj, norm_ggid) -> employee_name (from forecast file)

    def load(self, path):
        df_raw = fast_load_excel(path)
        self.items = [ExpenditureItem(row) for row in df_raw.to_dict("records")]

    def load_forecast(self, path):
        """Parse un fichier Forecast et remplit self.forecast_data.

        Detection dynamique de la structure :
        - Ligne d'en-tetes trouvee par scan des 20 premieres lignes
        - Colonnes d'identite (Project Num, GGID, …) trouvees par nom
        - Blocs mensuels localises via les noms de mois (FR/EN) et les
          sous-colonnes (Grade, F.RC, Vaca, Prod, Cost, FTE)
        - Blocs Total / TR Q* exclus automatiquement
        """
        import openpyxl
        from datetime import datetime

        # ---- Helpers locaux ----
        def _col_letter(idx_0):
            """0-indexed column number -> Excel letter(s)."""
            result = ""
            n = idx_0
            while True:
                result = chr(n % 26 + ord('A')) + result
                n = n // 26 - 1
                if n < 0:
                    break
            return result

        def _to_float(val):
            try:
                if val is None or str(val).strip() in ("", "-"):
                    return 0.0
                return float(val)
            except (ValueError, TypeError):
                return 0.0

        SCAN_MAX_COL = 200  # ne pas scanner les colonnes fantomes Excel

        # Marqueurs d'identite attendus dans la ligne d'en-tetes
        IDENTITY_MARKERS = {'project num', 'project name', 'ggid',
                            'employee name', 'grade'}

        # Mapping noms de mois FR + EN -> numero de mois (1-12)
        MONTH_NAMES_MAP = {}
        for _names, _num in [
            (['janvier', 'january', 'jan'], 1),
            (['fevrier', 'février', 'february', 'feb'], 2),
            (['mars', 'march', 'mar'], 3),
            (['avril', 'april', 'apr'], 4),
            (['mai', 'may'], 5),
            (['juin', 'june', 'jun'], 6),
            (['juillet', 'july', 'jul'], 7),
            (['août', 'aout', 'august', 'aug'], 8),
            (['septembre', 'september', 'sep', 'sept'], 9),
            (['octobre', 'october', 'oct'], 10),
            (['novembre', 'november', 'nov'], 11),
            (['decembre', 'décembre', 'december', 'dec'], 12),
        ]:
            for _n in _names:
                MONTH_NAMES_MAP[_n] = _num

        # Noms attendus dans un bloc mensuel (ordre canonique)
        EXPECTED_SUB_ORDER = ['grade', 'f.rc', 'vaca', 'prod', 'cost', 'fte']
        SUB_COL_SET = set(EXPECTED_SUB_ORDER)

        # Valeurs parasites : si Project Num == un de ces mots, c'est une
        # ligne mal identifiee (en-tete duplique, etc.), pas une donnee.
        PARASITE_VALUES = (
            IDENTITY_MARKERS
            | set(MONTH_NAMES_MAP.keys())
            | SUB_COL_SET
            | {'total'}
        )

        print("\n" + "=" * 70)
        print("[FORECAST IMPORT] Debut du chargement du fichier Forecast")
        print("=" * 70)

        # =================================================================
        # Ouverture en mode normal pour la detection de structure
        # (merged_cells, dimensions, …). Fichiers Forecast = petits,
        # donc pas de probleme de performance.
        # =================================================================
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active

        print(f"\n  Feuille active : '{ws.title}'")
        print(f"  Dimensions     : {ws.dimensions}")

        # =================================================================
        # 1. DETECTION DE LA LIGNE D'EN-TETES
        # =================================================================
        print("\n--- 1. DETECTION DE LA LIGNE D'EN-TETES ---")

        header_row_idx = None   # 1-indexed
        header_cells = {}       # col_0idx -> str value (stripped)

        for r in range(1, min(ws.max_row + 1, 21)):
            row_map = {}
            hits = 0
            for cell in next(ws.iter_rows(min_row=r, max_row=r,
                                          max_col=SCAN_MAX_COL,
                                          values_only=False)):
                if cell.value is not None and isinstance(cell.value, str):
                    cleaned = cell.value.strip()
                    if cleaned:
                        row_map[cell.column - 1] = cleaned
                        if cleaned.lower() in IDENTITY_MARKERS:
                            hits += 1
            if hits >= 3:
                header_row_idx = r
                header_cells = row_map
                break

        if header_row_idx is None:
            print("  [ERROR] Impossible de trouver la ligne d'en-tetes dans "
                  "les 20 premieres lignes.")
            print("  Marqueurs recherches : "
                  + ", ".join(sorted(IDENTITY_MARKERS)))
            print("  Import annule (aucune donnee chargee).")
            wb.close()
            return

        print(f"  Ligne d'en-tetes trouvee : ligne {header_row_idx}")
        print(f"  Colonnes avec texte dans cette ligne : {len(header_cells)}")

        # =================================================================
        # 2. DETECTION DYNAMIQUE DES COLONNES D'IDENTITE
        # =================================================================
        print("\n--- 2. DETECTION DES COLONNES D'IDENTITE ---")

        IDENTITY_LOOKUP = {
            'project num': 'project_num',
            'project name': 'project_name',
            'ggid': 'ggid',
            'employee name': 'employee_name',
        }

        identity_cols = {}  # key (ex: 'project_num') -> col_0idx

        for col_0, val in header_cells.items():
            key = val.lower()
            if key in IDENTITY_LOOKUP:
                field = IDENTITY_LOOKUP[key]
                if field not in identity_cols:
                    identity_cols[field] = col_0

        col_proj = identity_cols.get('project_num')
        col_proj_name = identity_cols.get('project_name')
        col_ggid = identity_cols.get('ggid')
        col_emp_name = identity_cols.get('employee_name')

        for label, idx in [('Project Num', col_proj),
                           ('Project Name', col_proj_name),
                           ('GGID', col_ggid),
                           ('Employee Name', col_emp_name)]:
            if idx is not None:
                print(f"  [OK] '{label}' -> colonne {_col_letter(idx)} "
                      f"(0-idx {idx})")
            else:
                print(f"  [WARNING] '{label}' introuvable dans la ligne "
                      f"d'en-tetes")

        if col_proj is None or col_ggid is None:
            print("  [ERROR] Colonnes obligatoires manquantes "
                  "(Project Num et/ou GGID). Import annule.")
            wb.close()
            return

        # =================================================================
        # 3. DETECTION DYNAMIQUE DES BLOCS MENSUELS
        # =================================================================
        print("\n--- 3. DETECTION DES BLOCS MENSUELS ---")

        # 3a. Scanner les lignes au-dessus du header (et le header lui-meme)
        #     pour reperer les labels de mois / Total / TR.
        block_labels = {}  # col_0idx -> ('month', num) | ('skip', label_text)

        scan_start = max(1, header_row_idx - 5)
        for r in range(scan_start, header_row_idx + 1):
            for cell in next(ws.iter_rows(min_row=r, max_row=r,
                                          max_col=SCAN_MAX_COL,
                                          values_only=False)):
                if cell.value is None or not isinstance(cell.value, str):
                    continue
                val_raw = cell.value.strip()
                val_lower = val_raw.lower()
                c0 = cell.column - 1

                if val_lower in MONTH_NAMES_MAP:
                    block_labels[c0] = ('month', MONTH_NAMES_MAP[val_lower])
                elif val_lower == 'total' or val_lower.startswith('total'):
                    block_labels[c0] = ('skip', val_raw)
                elif val_lower.startswith('tr '):
                    block_labels[c0] = ('skip', val_raw)

        # 3b. Pour chaque label de mois, identifier les sous-colonnes
        #     (Grade, F.RC, Vaca, Prod, Cost, FTE) dans le header.
        month_blocks = {}  # month_num -> {frc: c0, prod: c0, cost: c0}

        for label_col in sorted(block_labels.keys()):
            label_type, label_val = block_labels[label_col]

            if label_type == 'skip':
                print(f"  [SKIP] Bloc '{label_val}' a col "
                      f"{_col_letter(label_col)} (Total/TR)")
                continue

            month_num = label_val

            # Verifier les 6 colonnes a partir de label_col dans le header
            sub_cols = {}
            order_ok = True
            for offset, expected in enumerate(EXPECTED_SUB_ORDER):
                c = label_col + offset
                actual = header_cells.get(c, '').lower()
                if actual == expected:
                    sub_cols[expected] = c
                else:
                    order_ok = False

            if not order_ok:
                # Fallback : chercher chaque sous-colonne par nom dans
                # une plage elargie (label_col .. label_col+7)
                print(f"  [WARNING] Mois {month_num:02d} "
                      f"(col {_col_letter(label_col)}) : ordre des "
                      f"sous-colonnes non standard, recherche par nom")
                sub_cols = {}
                for c in range(label_col, label_col + 8):
                    name = header_cells.get(c, '').lower()
                    if name in SUB_COL_SET:
                        sub_cols[name] = c

            frc_col = sub_cols.get('f.rc')
            prod_col = sub_cols.get('prod')
            cost_col = sub_cols.get('cost')

            if frc_col is None or prod_col is None or cost_col is None:
                missing = [n for n, v in [('F.RC', frc_col),
                                          ('Prod', prod_col),
                                          ('Cost', cost_col)] if v is None]
                print(f"  [WARNING] Mois {month_num:02d} : colonnes manquantes "
                      f"{missing} — ce mois sera ignore")
                continue

            month_blocks[month_num] = {
                'frc': frc_col, 'prod': prod_col, 'cost': cost_col
            }
            print(f"  [OK] Mois {month_num:02d} : "
                  f"F.RC=col {_col_letter(frc_col)}, "
                  f"Prod=col {_col_letter(prod_col)}, "
                  f"Cost=col {_col_letter(cost_col)}")

        if not month_blocks:
            print("  [ERROR] Aucun bloc mensuel valide detecte. "
                  "Import annule.")
            wb.close()
            return

        print(f"  => {len(month_blocks)} bloc(s) mensuel(s) valide(s)")

        # =================================================================
        # 4. DETECTION DES BLOCS PROJETS
        #    Chercher les plages fusionnees dans la colonne Project Num
        #    detectee dynamiquement, ou fallback par rupture de valeur.
        # =================================================================
        print("\n--- 4. DETECTION DES BLOCS PROJETS ---")

        proj_col_1idx = col_proj + 1   # 1-indexed pour l'API merged_cells
        ggid_col_1idx = col_ggid + 1
        emp_col_1idx = (col_emp_name + 1) if col_emp_name is not None else None

        merged_ranges = sorted(
            [mr for mr in ws.merged_cells.ranges
             if mr.min_col == proj_col_1idx],
            key=lambda r: r.min_row,
        )

        if merged_ranges:
            print(f"  {len(merged_ranges)} plage(s) fusionnee(s) en colonne "
                  f"{_col_letter(col_proj)} :")
            for mr in merged_ranges:
                anchor = ws.cell(row=mr.min_row, column=proj_col_1idx).value
                n_rows = mr.max_row - mr.min_row + 1
                tot_row = mr.max_row + 1
                tot_ggid = ws.cell(row=tot_row, column=ggid_col_1idx).value
                tot_emp = (ws.cell(row=tot_row, column=emp_col_1idx).value
                           if emp_col_1idx else None)
                has_total = (tot_ggid is None
                             and (emp_col_1idx is None or tot_emp is None))
                print(
                    f"    Projet '{anchor}' : lignes {mr.min_row}-{mr.max_row} "
                    f"({n_rows} consultants), "
                    f"ligne Total={tot_row} "
                    f"{'[detectee et exclue]' if has_total else '[NON detectee — WARNING]'}"
                )
        else:
            print(f"  [INFO] Aucune plage fusionnee en colonne "
                  f"{_col_letter(col_proj)}.")
            print("  Fallback : detection par rupture de valeur :")
            proj_groups = {}
            for row_cells in ws.iter_rows(
                    min_row=header_row_idx + 1, min_col=1,
                    max_col=SCAN_MAX_COL, values_only=False):
                p_val = row_cells[col_proj].value if col_proj < len(row_cells) else None
                g_val = row_cells[col_ggid].value if col_ggid < len(row_cells) else None
                if p_val is not None and g_val is not None:
                    proj_groups.setdefault(
                        str(p_val).strip(), []).append(row_cells[0].row)
            if proj_groups:
                for pv, rows in proj_groups.items():
                    print(f"    Projet '{pv}' : {len(rows)} consultant(s) "
                          f"(lignes {rows[0]}-{rows[-1]})")
            else:
                print("    [WARNING] Aucun Project Num trouve.")

        wb.close()

        # =================================================================
        # PARSING PRINCIPAL (read_only pour la performance)
        # =================================================================
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.active
            current_year = datetime.now().year
            current_proj = None

            # Compteurs de diagnostic
            projects_found = set()
            consultants_by_project = {}
            rows_processed = 0
            rows_skipped = 0

            data_start_row = header_row_idx + 1

            for row in ws.iter_rows(min_row=data_start_row, values_only=True):
                # Acces securise aux colonnes detectees dynamiquement
                def _safe(idx):
                    if idx is not None and idx < len(row):
                        return row[idx]
                    return None

                proj_val = _safe(col_proj)
                ggid_val = _safe(col_ggid)

                # Mettre a jour le projet courant
                if proj_val is not None:
                    # Garde-fou : ignorer les lignes parasites
                    pv_lower = str(proj_val).strip().lower()
                    if pv_lower in PARASITE_VALUES:
                        continue
                    current_proj = proj_val

                # Ignorer les lignes Total (GGID vide)
                if ggid_val is None:
                    rows_skipped += 1
                    continue

                rows_processed += 1
                norm_proj = ExpenditureItem._normalize_key(current_proj)
                norm_ggid = str(ggid_val).strip().replace("6427_", "")

                # Stocker le nom du consultant depuis le forecast
                emp_name_val = _safe(col_emp_name)
                if emp_name_val is not None:
                    self.forecast_names[(norm_proj, norm_ggid)] = str(emp_name_val).strip()

                projects_found.add(norm_proj)
                consultants_by_project.setdefault(norm_proj, set()).add(
                    (str(ggid_val).strip(), norm_ggid)
                )

                for month_num, block in sorted(month_blocks.items()):
                    month_key = f"{current_year}-{month_num:02d}"

                    frc = _to_float(_safe(block['frc']))
                    prod = _to_float(_safe(block['prod']))
                    cost = _to_float(_safe(block['cost']))

                    if frc != 0 or prod != 0 or cost != 0:
                        self.forecast_data[
                            (norm_proj, norm_ggid, month_key)
                        ] = {
                            "ncr": frc,
                            "working_days": prod,
                            "bill_amount": cost,
                        }

            # =============================================================
            # 5. MAPPING AVEC LE FICHIER ICB
            # =============================================================
            print("\n--- 5. MAPPING AVEC LE FICHIER ICB ---")

            icb_projects = set()
            icb_employees_by_project = {}
            for item in self.filtered_items:
                p = ExpenditureItem._normalize_key(item.cc_project_number)
                e = str(item.employee_number).strip().replace("6427_", "")
                icb_projects.add(p)
                icb_employees_by_project.setdefault(p, set()).add(e)

            matched_projects = projects_found & icb_projects
            unmatched_projects = projects_found - icb_projects

            print(f"  Projets Forecast : {len(projects_found)}")
            print(f"  Projets ICB      : {len(icb_projects)}")
            print(f"  Projets matches  : "
                  f"{len(matched_projects)}/{len(projects_found)}")
            if matched_projects:
                for p in sorted(matched_projects):
                    print(f"    [MATCH]    {p}")
            if unmatched_projects:
                print(f"  [WARNING] {len(unmatched_projects)} projet(s) "
                      f"Forecast NON trouves dans l'ICB :")
                for p in sorted(unmatched_projects):
                    print(f"    [NO MATCH] {p}")

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
                            f"normalise='{norm_ggid}' dans projet "
                            f"'{proj_norm}' — absent de l'ICB"
                        )

            total_cons = total_matched_cons + total_unmatched_cons
            print(f"  Consultants matches : "
                  f"{total_matched_cons}/{total_cons}")
            if total_unmatched_cons:
                print(f"  [WARNING] {total_unmatched_cons} consultant(s) "
                      f"non matches")

            # =============================================================
            # 6. RESULTAT FINAL
            # =============================================================
            cells_filled = len(self.forecast_data) * 3

            print("\n--- 6. RESULTAT FINAL ---")
            print(f"  En-tetes detectes    : ligne {header_row_idx}")
            print(f"  Colonnes identite    : Project Num="
                  f"{_col_letter(col_proj)}, GGID="
                  f"{_col_letter(col_ggid)}")
            print(f"  Blocs mensuels       : {len(month_blocks)}")
            print(f"  Lignes consultants   : {rows_processed}")
            print(f"  Lignes ignorees      : {rows_skipped}")
            print(f"  Entrees forecast     : {len(self.forecast_data)}")
            print(f"  Cellules remplies    : {cells_filled} "
                  f"(NCR + Working Days + Bill Amount)")
            print(
                f"\n[FORECAST IMPORT] {len(projects_found)} projets, "
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
                    norm_proj = ExpenditureItem._normalize_key(cc['cc_project_number'])

                    # --- Col 5 : Forecast Bill Amount ---
                    # Distinguer "aucune donnee forecast" (vide) de "forecast = 0" (ecrire 0)
                    has_any_forecast = False
                    forecast_ba_total = 0.0
                    for fk, fd in self.forecast_data.items():
                        if fk[0] == norm_proj and fk[2] == month:
                            has_any_forecast = True
                            forecast_ba_total += fd["bill_amount"]

                    forecast_cell_value = round(forecast_ba_total, 2) if has_any_forecast else ""

                    # --- Col 8 : Comment auto-genere ---
                    comment = ""
                    icb_ba = cc['totals']['bill_amount']
                    gap_val = round(icb_ba - forecast_ba_total, 2) if has_any_forecast else 0

                    if has_any_forecast and abs(gap_val) > GAP_TOLERANCE_AMOUNT:
                        # Identifier les causes par consultant
                        # 1. Construire les sets ICB et Forecast (normalises)
                        icb_emps = {}  # norm_emp -> cons dict
                        for cons in cc['consultants']:
                            ne = str(cons['employee_id']).strip().replace("6427_", "")
                            icb_emps[ne] = cons

                        forecast_emps = set()
                        for fk in self.forecast_data:
                            if fk[0] == norm_proj and fk[2] == month:
                                forecast_emps.add(fk[1])

                        icb_set = set(icb_emps.keys())

                        # 2. Consultants absents
                        only_icb = icb_set - forecast_emps      # dans ICB, pas dans Forecast
                        only_forecast = forecast_emps - icb_set  # dans Forecast, pas dans ICB

                        # Helper : formater un identifiant consultant avec son nom
                        def _id_name(ne, source='icb'):
                            """Retourne 'ID (Nom)' ou juste 'ID' si nom inconnu."""
                            name = None
                            if source == 'icb' and ne in icb_emps:
                                name = icb_emps[ne].get('name')
                            if not name:
                                name = self.forecast_names.get((norm_proj, ne))
                            if name:
                                return f"{ne} ({name})"
                            return ne

                        # Helper : limiter la liste a 5 elements max
                        def _limit_list(entries, max_show=5):
                            if len(entries) <= max_show:
                                return ", ".join(entries)
                            shown = ", ".join(entries[:max_show])
                            return f"{shown} ... et {len(entries) - max_show} autre(s)"

                        # 3. Pour les consultants communs, detecter ecarts NCR / WD
                        common = icb_set & forecast_emps
                        details_wd = []   # [(id_name_str, wd_icb, wd_fc), ...]
                        details_ncr = []  # [(id_name_str, ncr_icb, ncr_fc), ...]
                        for ne in sorted(common):
                            cons = icb_emps[ne]
                            fk = (norm_proj, ne, month)
                            fd = self.forecast_data.get(fk)
                            if not fd:
                                continue
                            label = _id_name(ne, 'icb')
                            # Comparaison WD
                            wd_icb = round(cons['working_days'], 2)
                            wd_fc = round(fd['working_days'], 2)
                            if wd_icb and wd_fc and abs(wd_icb - wd_fc) > GAP_TOLERANCE_DAYS:
                                details_wd.append(f"{label} ICB={wd_icb}j/Forecast={wd_fc}j")
                            # Comparaison NCR
                            ncr_icb = round(cons['ncr'], 2)
                            ncr_fc = round(fd['ncr'], 2)
                            if ncr_icb and ncr_fc and abs(ncr_icb - ncr_fc) > GAP_TOLERANCE_AMOUNT:
                                details_ncr.append(f"{label} ICB={ncr_icb}/Forecast={ncr_fc}")

                        # 4. Composer le commentaire
                        parts = []

                        if details_wd:
                            parts.append(f"Ecart lie aux jours travailles : {_limit_list(details_wd)}")
                        if details_ncr:
                            parts.append(f"Ecart lie au taux (NCR) : {_limit_list(details_ncr)}")

                        if only_forecast:
                            entries = [_id_name(ne, 'forecast') for ne in sorted(only_forecast)]
                            parts.append(f"Consultant(s) absent(s) de l'ICB : {_limit_list(entries)}")

                        if only_icb:
                            entries = [_id_name(ne, 'icb') for ne in sorted(only_icb)]
                            parts.append(f"Consultant(s) absent(s) du Forecast : {_limit_list(entries)}")

                        if not parts:
                            parts.append("Ecart non explique par jours travailles ou taux — a verifier manuellement")

                        comment = " ; ".join(parts)

                    ws.append([
                        cc['project_name'] or "",
                        cc['cc_project_number'],
                        cc['totals']['working_days'],
                        cc['totals']['bill_amount'],
                        forecast_cell_value,       # Forecast bill amount
                        f"=D{current_row}-E{current_row}", # Gap (formule Excel)
                        cc['totals']['raw_cost'],
                        comment                    # Comment auto-genere
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
                ws.column_dimensions['H'].width = 60

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

    # =====================================================================
    # DIAGNOSTIC TEMPORAIRE — a retirer apres resolution du bug
    # =====================================================================
    print("\n" + "=" * 70)
    print("[DIAG TRANSACTIONS] Clic sur ligne consultant")
    print("=" * 70)

    # 1. Employee Number BRUT (tel que recu par l'appel API)
    print(f"  1. Employee Number BRUT recu par l'API : '{emp}' (type={type(emp).__name__})")
    print(f"     CC Project Number BRUT recu par l'API : '{cc_proj}' (type={type(cc_proj).__name__})")

    # 2. Employee Number NORMALISE (applique _normalize_key pour comparaison)
    norm_emp = ExpenditureItem._normalize_key(emp)
    norm_cc = ExpenditureItem._normalize_key(cc_proj)
    print(f"  2. Employee Number NORMALISE : '{norm_emp}'")
    print(f"     CC Project Number NORMALISE : '{norm_cc}'")

    # 3. Nombre d'expenditure items AVANT filtres de dates
    #    On cherche dans TOUS les items filtres (sans filtre de mois/date)
    all_items_no_filter = pipeline.filtered_items
    items_before_filter = [
        i for i in all_items_no_filter
        if i.employee_number == emp and i.cc_project_number == cc_proj
    ]
    # Aussi tester avec comparaison normalisee
    items_before_filter_normalized = [
        i for i in all_items_no_filter
        if ExpenditureItem._normalize_key(i.employee_number) == norm_emp
        and i.cc_project_number == norm_cc
    ]
    print(f"  3. Items AVANT filtres de dates (comparaison directe) : {len(items_before_filter)}")
    print(f"     Items AVANT filtres de dates (comparaison normalisee) : {len(items_before_filter_normalized)}")
    if len(items_before_filter) != len(items_before_filter_normalized):
        print(f"     *** DIVERGENCE DETECTEE ! La comparaison directe et normalisee donnent des resultats differents ***")
        print(f"     Cela indique un probleme de type ou de format entre la valeur API et les donnees source.")

    # 4. Filtres de dates actuellement actifs
    print(f"  4. Filtres actifs : month='{month}', start_date='{start_date}', end_date='{end_date}'")

    # 5. Items APRES application des filtres
    items = pipeline._get_items(month_filter=month, start_date=start_date, end_date=end_date)
    txns = [
        i for i in items
        if i.employee_number == emp and i.cc_project_number == cc_proj
    ]
    txns_normalized = [
        i for i in items
        if ExpenditureItem._normalize_key(i.employee_number) == norm_emp
        and i.cc_project_number == norm_cc
    ]
    print(f"  5. Items APRES filtres (comparaison directe) : {len(txns)}")
    print(f"     Items APRES filtres (comparaison normalisee) : {len(txns_normalized)}")

    # 6. Si 0 items trouves en comparaison directe : inspecter les donnees source
    if len(items_before_filter) == 0:
        print(f"  6. *** 0 ITEMS TROUVES — Inspection des donnees source ***")
        # Lister les employee_numbers presents pour ce cc_project
        emps_for_proj = set()
        emps_for_proj_raw = {}  # norm -> (raw_value, raw_type)
        for i in all_items_no_filter:
            if i.cc_project_number == cc_proj or i.cc_project_number == norm_cc:
                raw_e = i.employee_number
                norm_e = ExpenditureItem._normalize_key(raw_e)
                emps_for_proj.add(norm_e)
                emps_for_proj_raw[norm_e] = (raw_e, type(raw_e).__name__)
        
        print(f"     CC Project '{cc_proj}' : {len(emps_for_proj)} employee(s) distincts dans les donnees source")
        
        # Chercher si l'employee existe sous un format different
        if norm_emp in emps_for_proj:
            raw_val, raw_type = emps_for_proj_raw[norm_emp]
            print(f"     -> Employee TROUVE sous format different !")
            print(f"        Valeur dans les donnees : '{raw_val}' (type={raw_type})")
            print(f"        Valeur recue par l'API  : '{emp}' (type=str)")
            print(f"        => Le bug est un PROBLEME DE TYPE (la comparaison directe echoue)")
        else:
            print(f"     -> Employee '{norm_emp}' est ABSENT de ce projet dans les donnees source")
            # Montrer les employees presents pour aider au debug
            sample = list(emps_for_proj)[:10]
            print(f"     -> Employees presents (max 10) : {sample}")
            # Aussi verifier si l'employee existe dans un AUTRE projet
            other_projects = set()
            for i in all_items_no_filter:
                if ExpenditureItem._normalize_key(i.employee_number) == norm_emp:
                    other_projects.add(i.cc_project_number)
            if other_projects:
                print(f"     -> Cet employee existe dans d'autres projets : {other_projects}")
            else:
                print(f"     -> Cet employee n'existe dans AUCUN projet des donnees source")
    else:
        print(f"  6. (Pas d'inspection necessaire — {len(items_before_filter)} items trouves)")

    print("=" * 70 + "\n")
    # =====================================================================
    # FIN DIAGNOSTIC TEMPORAIRE
    # =====================================================================

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