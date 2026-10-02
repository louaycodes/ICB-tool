"""
Application ICB - Capgemini
Backend Flask : charge le fichier Excel, filtre, groupe, et expose les donnees
via des endpoints JSON pour l'interface web.
"""

import os
import sys
import io
import math
import tempfile
import pandas as pd
from datetime import datetime, date, timedelta
from collections import OrderedDict
from flask import Flask, request, jsonify, render_template
from openpyxl import load_workbook

def get_resource_path(relative_path):
    """ Obtient le chemin absolu vers la ressource, fonctionne en dev et pour PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Initialisation de Flask avec les chemins gérés par PyInstaller
app = Flask(
    __name__,
    template_folder=get_resource_path('templates'),
    static_folder=get_resource_path('static')
)

# Si vous accédez au dossier uploads dans votre code, utilisez aussi :
UPLOAD_FOLDER = get_resource_path('uploads')

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

# Les seuils de tolerance sont desormais envoyes dynamiquement depuis le frontend.

# --- Flask app ---
app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 Mo max pour l'upload

# stockage en memoire du pipeline (une seule session)
pipeline = None


# ---------------------------------------------------------------------------
# Lecture rapide de gros fichiers Excel
# ---------------------------------------------------------------------------
def fast_load_excel(path, sheet_name=0, dynamic_header=False):
    """Charge un fichier Excel en DataFrame.

    Si *dynamic_header* est True, la fonction scanne les 20 premieres lignes
    pour trouver celle contenant les en-tetes attendus (meme logique que le
    chargement Forecast). Sinon, la ligne 1 est utilisee comme en-tete.

    Retourne un DataFrame pandas.
    """
    wb = load_workbook(path, read_only=True, data_only=True)

    # --- Selection de la feuille ---
    if isinstance(sheet_name, str):
        if sheet_name not in wb.sheetnames:
            available = wb.sheetnames
            wb.close()
            raise ValueError(
                f"Feuille '{sheet_name}' introuvable dans le classeur. "
                f"Feuilles disponibles : {available}"
            )
        ws = wb[sheet_name]
    else:
        ws = wb.worksheets[sheet_name]

    SCAN_MAX_COL = 200

    if dynamic_header:
        # --- Scan dynamique des en-tetes ---
        # Marqueurs attendus pour identifier la ligne d'en-tetes ICB
        ICB_HEADER_MARKERS = {
            'employee number', 'cc project number', 'project name',
            'expenditure type', 'quantity',
        }

        header_row_idx = None   # 1-indexed
        header_cells = {}       # col_0idx -> str value
        min_data_col = None     # 0-indexed : premiere colonne non-vide

        for r in range(1, min(ws.max_row + 1 if ws.max_row else 2, 21)):
            row_data = next(ws.iter_rows(min_row=r, max_row=r,
                                         min_col=1, max_col=SCAN_MAX_COL,
                                         values_only=True))
            row_map = {}
            hits = 0
            for ci, val in enumerate(row_data):
                if val is not None and isinstance(val, str) and val.strip():
                    row_map[ci] = val.strip()
                    if val.strip().lower() in ICB_HEADER_MARKERS:
                        hits += 1
            if hits >= 3:
                header_row_idx = r
                header_cells = row_map
                break

        if header_row_idx is None:
            wb.close()
            raise ValueError(
                "Impossible de trouver la ligne d'en-tetes dans les 20 "
                "premieres lignes de la feuille. Marqueurs recherches : "
                + ", ".join(sorted(ICB_HEADER_MARKERS))
            )

        # Determiner la premiere et la derniere colonne contenant un en-tete
        col_indices = sorted(header_cells.keys())
        min_data_col = col_indices[0]       # 0-indexed
        max_data_col = col_indices[-1]       # 0-indexed
        ncols = max_data_col - min_data_col + 1

        # Construire le header (peut contenir des None pour des trous)
        header = []
        for ci in range(min_data_col, max_data_col + 1):
            header.append(header_cells.get(ci))

        print(f"[ICB] En-tetes detectes a la ligne {header_row_idx}, "
              f"colonnes {min_data_col}-{max_data_col} ({ncols} colonnes)")

        # Lire les donnees a partir de la ligne suivante
        rows_iter = ws.iter_rows(
            min_row=header_row_idx + 1,
            min_col=min_data_col + 1,   # 1-indexed pour openpyxl
            max_col=max_data_col + 1,   # 1-indexed
            values_only=True,
        )
        data = [r for r in rows_iter if any(v is not None for v in r)]

    else:
        # --- Mode classique : en-tetes en ligne 1, donnees en col A ---
        header_row = next(ws.iter_rows(min_row=1, max_row=1, min_col=1,
                                        max_col=SCAN_MAX_COL, values_only=True))
        ncols = max((i for i, v in enumerate(header_row) if v is not None),
                    default=-1) + 1
        if ncols == 0:
            wb.close()
            raise ValueError("Feuille vide ou en-tete introuvable.")
        header = list(header_row[:ncols])
        rows_iter = ws.iter_rows(min_row=2, min_col=1, max_col=ncols,
                                  values_only=True)
        data = [r for r in rows_iter if any(v is not None for v in r)]

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
        def _clean(val):
            if pd.isna(val):
                return None
            if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
                return None
            return val

        self.item_id = _clean(row.get("Expenditure Item Id"))
        self.project_name = _clean(row.get("Project Name")) or ""
        self.cc_project_number = self._normalize_key(_clean(row.get("CC Project Number")))
        self.task_number = _clean(row.get("Task Number"))
        self.employee_number = self._normalize_key(_clean(row.get("Employee Number")))
        self.full_name = _clean(row.get("Full Name"))
        self.emp_grade = _clean(row.get("Emp Grade"))
        self.expenditure_type = _clean(row.get("Expenditure Type"))
        self.uom = _clean(row.get("UOM"))
        self.quantity = self._safe_float(_clean(row.get("Quantity")))
        br = self._safe_float(_clean(row.get("Bill Rate")))
        self.bill_rate = br * 8.0 if br is not None else None
        self.raw_cost = self._safe_float(_clean(row.get("Raw Cost")))
        self.bill_amount = self._safe_float(_clean(row.get("Project Transfer Price")))
        self.invoice_type = _clean(row.get("Invoice Type"))
        self.date = excel_serial_to_date(_clean(row.get("Expenditure Item Date")))
        # Colonne de correlation pour la detection d'anomalie
        raw_corr = _clean(row.get("FULL NCR to Direct NCR Correlation"))
        self.ncr_correlation = self._safe_float(raw_corr)

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
    def _safe_float(value):
        """Convertit une valeur en float de maniere securisee.
        Retourne None si la valeur est None ou non convertible.
        Gere le cas ou Excel renvoie des nombres sous forme de string."""
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        # Tenter la conversion depuis une string
        try:
            f = float(value)
            if math.isnan(f) or math.isinf(f):
                return None
            return f
        except (ValueError, TypeError):
            return None

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
        # Charger la feuille "ICB_DB" avec detection dynamique des en-tetes
        df_raw = fast_load_excel(path, sheet_name="DB_ICB", dynamic_header=True)

        # --- Detection dynamique de la colonne "Project Transfer Price" ---
        # Accepte 3 variantes ordonnees de noms possibles :
        candidate_names = [
            "Project Transfer Price",
            "Project Revenue Price",
            "Revenue Project TP",
        ]
        ptp_col = None
        for candidate in candidate_names:
            c_norm = candidate.strip().lower()
            for col in df_raw.columns:
                if col is not None and str(col).strip().lower() == c_norm:
                    ptp_col = col
                    break
            if ptp_col is not None:
                if candidate != candidate_names[0]:
                    print(f"[INFO] Colonne alternative '{ptp_col}' detectee "
                          f"(utilisee a la place de '{candidate_names[0]}')")
                break

        if ptp_col is None:
            names_str = ", ".join(f"'{name}'" for name in candidate_names)
            print(f"[ERROR] Aucune des colonnes attendues ({names_str}) "
                  "n'a ete trouvee dans le fichier ICB.")
            print(f"  Colonnes disponibles : {list(df_raw.columns)}")
        elif str(ptp_col).strip() != "Project Transfer Price":
            # Renommer vers le nom canonique attendu par ExpenditureItem
            df_raw = df_raw.rename(columns={ptp_col: "Project Transfer Price"})
            print(f"[INFO] Colonne '{ptp_col}' renommee en 'Project Transfer Price'")

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
                f = float(val)
                import math
                if math.isnan(f) or math.isinf(f):
                    return 0.0
                return f
            except (ValueError, TypeError):
                return 0.0

        SCAN_MAX_COL = 200  # ne pas scanner les colonnes fantomes Excel

        # Marqueurs d'identite attendus dans la ligne d'en-tetes
        IDENTITY_MARKERS = {'project num', 'project name', 'icb id',
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
            'icb id': 'icb_id',
            'employee name': 'employee_name',
        }

        identity_cols = {}  # key (ex: 'project_num') -> col_0idx

        for col_0, val in header_cells.items():
            key = val.lower()
            if key in IDENTITY_LOOKUP:
                field = IDENTITY_LOOKUP[key]
                if field not in identity_cols:
                    identity_cols[field] = col_0

        # --- Detection speciale pour ICB ID (peut etre dans une colonne cachee) ---
        if 'icb_id' not in identity_cols:
            print("  [INFO] 'ICB ID' non trouvee dans les colonnes visibles, "
                  "recherche dans les colonnes cachees...")
            for cell in next(ws.iter_rows(min_row=header_row_idx,
                                          max_row=header_row_idx,
                                          max_col=SCAN_MAX_COL,
                                          values_only=False)):
                if cell.value is not None and isinstance(cell.value, str):
                    if cell.value.strip().lower() == 'icb id':
                        c0 = cell.column - 1
                        col_letter = _col_letter(c0)
                        # Verifier si la colonne est cachee
                        is_hidden = False
                        try:
                            dim = ws.column_dimensions.get(col_letter)
                            if dim and dim.hidden:
                                is_hidden = True
                        except Exception:
                            pass
                        identity_cols['icb_id'] = c0
                        print(f"  [OK] 'ICB ID' trouvee en colonne {col_letter} "
                              f"(0-idx {c0})"
                              f"{' [COLONNE CACHEE]' if is_hidden else ''}")
                        break

        col_proj = identity_cols.get('project_num')
        col_proj_name = identity_cols.get('project_name')
        col_icb_id = identity_cols.get('icb_id')
        col_emp_name = identity_cols.get('employee_name')

        for label, idx in [('Project Num', col_proj),
                           ('Project Name', col_proj_name),
                           ('ICB ID', col_icb_id),
                           ('Employee Name', col_emp_name)]:
            if idx is not None:
                print(f"  [OK] '{label}' -> colonne {_col_letter(idx)} "
                      f"(0-idx {idx})")
            else:
                print(f"  [WARNING] '{label}' introuvable dans la ligne "
                      f"d'en-tetes")

        if col_proj is None or col_icb_id is None:
            print("  [ERROR] Colonnes obligatoires manquantes "
                  "(Project Num et/ou ICB ID). Import annule.")
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
        icb_id_col_1idx = col_icb_id + 1
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
                tot_icb_id = ws.cell(row=tot_row, column=icb_id_col_1idx).value
                tot_emp = (ws.cell(row=tot_row, column=emp_col_1idx).value
                           if emp_col_1idx else None)
                has_total = (tot_icb_id is None
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
                g_val = row_cells[col_icb_id].value if col_icb_id < len(row_cells) else None
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
                icb_id_val = _safe(col_icb_id)

                # Mettre a jour le projet courant
                if proj_val is not None:
                    # Garde-fou : ignorer les lignes parasites
                    pv_lower = str(proj_val).strip().lower()
                    if pv_lower in PARASITE_VALUES:
                        continue
                    current_proj = proj_val

                # Ignorer les lignes Total (ICB ID vide)
                if icb_id_val is None:
                    rows_skipped += 1
                    continue

                rows_processed += 1
                norm_proj = ExpenditureItem._normalize_key(current_proj)
                norm_icb_id = str(icb_id_val).strip().replace("6427_", "")

                # Stocker le nom du consultant depuis le forecast
                emp_name_val = _safe(col_emp_name)
                if emp_name_val is not None:
                    self.forecast_names[(norm_proj, norm_icb_id)] = str(emp_name_val).strip()

                projects_found.add(norm_proj)
                consultants_by_project.setdefault(norm_proj, set()).add(
                    (str(icb_id_val).strip(), norm_icb_id)
                )

                for month_num, block in sorted(month_blocks.items()):
                    month_key = f"{current_year}-{month_num:02d}"

                    frc = _to_float(_safe(block['frc']))
                    prod = _to_float(_safe(block['prod']))
                    cost = _to_float(_safe(block['cost']))

                    # Ignorer si toutes les valeurs sont a 0 ou vides
                    if abs(frc) > 1e-6 or abs(prod) > 1e-6 or abs(cost) > 1e-6:
                        self.forecast_data[
                            (norm_proj, norm_icb_id, month_key)
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
                for raw_id, norm_id in sorted(cons_set):
                    if norm_id in icb_emps:
                        total_matched_cons += 1
                    else:
                        total_unmatched_cons += 1
                        print(
                            f"  [NO MATCH] Consultant brut='{raw_id}', "
                            f"normalise='{norm_id}' dans projet "
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
                  f"{_col_letter(col_proj)}, ICB ID="
                  f"{_col_letter(col_icb_id)}")
            print(f"  Blocs mensuels       : {len(month_blocks)}")
            print(f"  Lignes consultants   : {rows_processed}")
            print(f"  Lignes ignorees      : {rows_skipped}")
            print(f"  Entrees forecast     : {len(self.forecast_data)}")
            print(f"  Cellules remplies    : {cells_filled} "
                  f"(NCR + Working Days + Cost)")
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
        """Returns the CC Project -> Consultants hierarchy as a serializable dict.

        Uses a UNION strategy: all ICB expenditure items are kept, and
        forecast-only entries (consultant present in forecast but missing
        from ICB) are injected as additional rows so nothing is lost.
        Each consultant entry carries a 'source' field:
          - 'both'          : present in ICB and forecast
          - 'icb_only'      : present in ICB but no matching forecast entry
          - 'forecast_only' : present in forecast but missing from ICB
        """
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
                        "forecast_working_days": 0.0, "forecast_bill_amount": 0.0, "forecast_ncr": 0.0, "ncr": 0.0
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
                    "source": "icb_only",  # will be upgraded to "both" if forecast match found
                }

            cons = cc_data["consultants"][emp]
            cons["ncr"] = item.bill_rate or cons["ncr"]
            cons["working_days"] += (item.quantity or 0) / 8.0
            cons["total_bill_amount"] += item.bill_amount or 0
            cons["total_raw_cost"] += item.raw_cost or 0
            # A consultant is flagged as anomalous if at least one of their lines is
            if item.has_anomaly:
                cons["has_anomaly"] = True

            cc_data["totals"]["working_days"] += (item.quantity or 0) / 8.0
            cc_data["totals"]["raw_cost"] += item.raw_cost or 0
            cc_data["totals"]["bill_amount"] += item.bill_amount or 0

        # Determine active months for forecast enrichment
        active_months = []
        if month_filter:
            active_months = [month_filter]
        else:
            if start_date and end_date:
                start = datetime.strptime(start_date, "%Y-%m-%d").date()
                end = datetime.strptime(end_date, "%Y-%m-%d").date()
                for m in self.available_months:
                    m_date = datetime.strptime(m + "-01", "%Y-%m-%d").date()
                    if start <= m_date <= end:
                        active_months.append(m)
            else:
                active_months = self.available_months

        # Also collect all months present in forecast_data for the UNION
        # (covers months that exist in forecast but not in ICB)
        forecast_months_set = set()
        for fk in self.forecast_data:
            forecast_months_set.add(fk[2])
        if month_filter:
            forecast_active = [month_filter] if month_filter in forecast_months_set else []
        elif start_date and end_date:
            s = datetime.strptime(start_date, "%Y-%m-%d").date()
            e = datetime.strptime(end_date, "%Y-%m-%d").date()
            forecast_active = []
            for m in sorted(forecast_months_set):
                try:
                    m_date = datetime.strptime(m + "-01", "%Y-%m-%d").date()
                    if s <= m_date <= e:
                        forecast_active.append(m)
                except ValueError:
                    pass
        else:
            forecast_active = sorted(forecast_months_set)

        # Merge: use the broadest set of months for enrichment
        all_active_months = sorted(set(active_months) | set(forecast_active))

        # Enrich existing ICB consultants with forecast data
        for cc_proj, cc_data in hierarchy.items():
            for emp, cons in cc_data["consultants"].items():
                norm_proj = ExpenditureItem._normalize_key(cc_proj)
                norm_emp = str(emp).strip().replace("6427_", "")

                sum_frc = 0.0
                sum_prod = 0.0
                sum_cost = 0.0
                frc_count = 0
                has_forecast = False
                for m in all_active_months:
                    key = (norm_proj, norm_emp, m)
                    if key in self.forecast_data:
                        has_forecast = True
                        fd = self.forecast_data[key]
                        sum_frc += fd["ncr"]
                        sum_prod += fd["working_days"]
                        sum_cost += fd["bill_amount"]
                        if fd["ncr"] != 0:
                            frc_count += 1

                cons["forecast_working_days"] = sum_prod
                cons["forecast_bill_amount"] = sum_cost
                if frc_count > 0:
                    cons["forecast_ncr"] = sum_frc / frc_count
                elif sum_prod > 0:
                    cons["forecast_ncr"] = (sum_cost / sum_prod) * 8.0
                else:
                    cons["forecast_ncr"] = 0.0

                # Update source indicator
                if has_forecast:
                    cons["source"] = "both"
                # else stays "icb_only"

                cc_data["totals"]["forecast_working_days"] += cons["forecast_working_days"]
                cc_data["totals"]["forecast_bill_amount"] += cons["forecast_bill_amount"]

            # Average NCR for the project (informational)
            num_cons = len(cc_data["consultants"])
            if num_cons > 0:
                cc_data["totals"]["ncr"] = sum(c["ncr"] for c in cc_data["consultants"].values()) / num_cons
                cc_data["totals"]["forecast_ncr"] = sum(c["forecast_ncr"] for c in cc_data["consultants"].values()) / num_cons
            else:
                cc_data["totals"]["ncr"] = 0.0
                cc_data["totals"]["forecast_ncr"] = 0.0

        # ---------------------------------------------------------------
        # UNION STEP: inject forecast-only consultants into the hierarchy
        # These are entries in forecast_data whose (project, consultant)
        # combination has no corresponding ICB expenditure items.
        # ---------------------------------------------------------------
        if self.forecast_data:
            # Build a set of (norm_proj, norm_emp) already in hierarchy
            icb_keys = set()
            for cc_proj, cc_data in hierarchy.items():
                norm_proj = ExpenditureItem._normalize_key(cc_proj)
                for emp in cc_data["consultants"]:
                    norm_emp = str(emp).strip().replace("6427_", "")
                    icb_keys.add((norm_proj, norm_emp))

            # Collect forecast-only entries grouped by (project, consultant)
            forecast_only = {}  # (norm_proj, norm_emp) -> list of (month, fd)
            for fk, fd in self.forecast_data.items():
                f_proj, f_emp, f_month = fk
                if f_month not in all_active_months:
                    continue
                if (f_proj, f_emp) not in icb_keys:
                    # Ignorer les entrees dont les valeurs de travail/cout sont a 0 / vides
                    if abs(fd.get("working_days", 0)) < 1e-6 and abs(fd.get("bill_amount", 0)) < 1e-6:
                        continue
                    forecast_only.setdefault((f_proj, f_emp), []).append((f_month, fd))

            for (f_proj, f_emp), month_entries in forecast_only.items():
                # Aggregate forecast values for this consultant
                sum_frc = 0.0
                sum_prod = 0.0
                sum_cost = 0.0
                frc_count = 0
                for _m, fd in month_entries:
                    sum_frc += fd["ncr"]
                    sum_prod += fd["working_days"]
                    sum_cost += fd["bill_amount"]
                    if fd["ncr"] != 0:
                        frc_count += 1

                # Ne pas injecter si toutes les valeurs d'activite sont nulles
                if abs(sum_prod) < 1e-6 and abs(sum_cost) < 1e-6:
                    continue

                # Find or create the project node in hierarchy
                # We need the original cc_proj key; try to find a match
                target_cc = None
                for cc_proj in hierarchy:
                    if ExpenditureItem._normalize_key(cc_proj) == f_proj:
                        target_cc = cc_proj
                        break

                if target_cc is None:
                    # Project exists only in forecast — create a new project node
                    target_cc = f_proj
                    hierarchy[target_cc] = {
                        "project_name": "",  # unknown from forecast
                        "consultants": OrderedDict(),
                        "totals": {
                            "working_days": 0.0, "raw_cost": 0.0, "bill_amount": 0.0,
                            "forecast_working_days": 0.0, "forecast_bill_amount": 0.0,
                            "forecast_ncr": 0.0, "ncr": 0.0
                        },
                    }

                cc_data = hierarchy[target_cc]

                fc_ncr = 0.0
                if frc_count > 0:
                    fc_ncr = sum_frc / frc_count
                elif sum_prod > 0:
                    fc_ncr = (sum_cost / sum_prod) * 8.0

                # Resolve consultant name from forecast_names
                emp_name = self.forecast_names.get((f_proj, f_emp), "")

                cc_data["consultants"][f_emp] = {
                    "name": emp_name,
                    "grade": "",
                    "ncr": 0.0,
                    "working_days": 0.0,
                    "total_bill_amount": 0.0,
                    "total_raw_cost": 0.0,
                    "has_anomaly": False,
                    "forecast_working_days": sum_prod,
                    "forecast_bill_amount": sum_cost,
                    "forecast_ncr": fc_ncr,
                    "source": "forecast_only",
                }

                cc_data["totals"]["forecast_working_days"] += sum_prod
                cc_data["totals"]["forecast_bill_amount"] += sum_cost

            # Recompute average NCR for projects (uniquement sur les consultants actifs)
            for cc_proj, cc_data in hierarchy.items():
                active_cons = [c for c in cc_data["consultants"].values()
                               if not (c["working_days"] == 0 and c["total_bill_amount"] == 0 and c["forecast_working_days"] == 0 and c["forecast_bill_amount"] == 0)]
                num_cons = len(active_cons)
                if num_cons > 0:
                    cc_data["totals"]["ncr"] = sum(c["ncr"] for c in active_cons) / num_cons
                    cc_data["totals"]["forecast_ncr"] = sum(c["forecast_ncr"] for c in active_cons) / num_cons
                else:
                    cc_data["totals"]["ncr"] = 0.0
                    cc_data["totals"]["forecast_ncr"] = 0.0

        # Sort and serialize
        result = []
        for cc_proj in sorted(hierarchy.keys()):
            cc = hierarchy[cc_proj]
            consultants_list = []
            for emp_id, cons in cc["consultants"].items():
                if cons["working_days"] == 0 and cons["total_bill_amount"] == 0 and cons["forecast_working_days"] == 0 and cons["forecast_bill_amount"] == 0:
                    continue
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
                    "source": cons.get("source", "both"),
                })

            # Si le projet n'a aucun consultant actif et aucune depense ni forecast, ne pas l'afficher
            if not consultants_list and cc["totals"]["working_days"] == 0 and cc["totals"]["bill_amount"] == 0 and cc["totals"]["forecast_bill_amount"] == 0:
                continue

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

    def export_to_excel(self, month_filter, start_date, end_date, gap_tolerance=2.0, include_comments=True):
        import io
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.worksheet.table import Table, TableStyleInfo
        from openpyxl.utils import get_column_letter

        # 1. Obtenir les items dans l'intervalle pour determiner les mois concernes
        items_in_range = self._get_items(month_filter=month_filter, start_date=start_date, end_date=end_date)
        months_in_range = sorted(list(set(i.month_key for i in items_in_range)))

        wb = Workbook()
        # Supprimer la feuille par defaut si on va creer des feuilles specifiques
        default_ws = wb.active

        # Styles
        right_align = Alignment(horizontal="right")
        wrap_align = Alignment(wrap_text=True)

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

                # Column headers — "Generated comment" is optional
                headers = [
                    "EM_Lead", 
                    "Project Name", 
                    "CC Project Number", 
                    "Working Days", 
                    "ICB Transfer Price", 
                    "Forecast Transfer Price", 
                    "Gap", 
                    "Raw Cost", 
                    "Issue Type", 
                ]
                if include_comments:
                    headers.append("Generated comment")
                num_cols = len(headers)
                ws.append(headers)

                current_row = 2
                for cc in data:
                    norm_proj = ExpenditureItem._normalize_key(cc['cc_project_number'])

                    # --- Col 6 : Forecast Bill Amount ---
                    has_any_forecast = False
                    forecast_ba_total = 0.0
                    for fk, fd in self.forecast_data.items():
                        if fk[0] == norm_proj and fk[2] == month:
                            if abs(fd.get("working_days", 0)) > 1e-6 or abs(fd.get("bill_amount", 0)) > 1e-6:
                                has_any_forecast = True
                                forecast_ba_total += fd["bill_amount"]

                    forecast_cell_value = round(forecast_ba_total, 2) if has_any_forecast else ""

                    # --- Generated comment (English templates) ---
                    comment = ""
                    icb_ba = cc['totals']['bill_amount']
                    gap_val = round(icb_ba - forecast_ba_total, 2) if has_any_forecast else 0

                    if include_comments and has_any_forecast and abs(gap_val) > gap_tolerance:
                        icb_emps = {}
                        for cons in cc['consultants']:
                            ne = str(cons['employee_id']).strip().replace("6427_", "")
                            icb_emps[ne] = cons

                        def _id_name(ne):
                            name = icb_emps[ne].get('name')
                            if not name:
                                name = self.forecast_names.get((norm_proj, ne))
                            if name:
                                name_str = str(name).strip()
                                if not (name_str.startswith("(") and name_str.endswith(")")):
                                    return f"({name_str})"
                                return name_str
                            return f"({ne})" if ne else "(Unknown)"

                        def _limit_list(entries):
                            formatted_entries = [f"  - {e}" for e in entries]
                            return "\n".join(formatted_entries)

                        details_wd = []
                        details_ncr = []
                        details_ba = []

                        for ne in sorted(icb_emps.keys()):
                            cons = icb_emps[ne]
                            fk = (norm_proj, ne, month)
                            fd = self.forecast_data.get(fk)
                            if not fd:
                                continue

                            label = _id_name(ne)

                            # Working days comparison: strict equality
                            wd_icb = round(cons['working_days'], 2)
                            wd_fc = round(fd['working_days'], 2)
                            if wd_icb != wd_fc:
                                details_wd.append(f"{label} : ICB = {wd_icb} D / Forecast = {wd_fc} D")

                            # NCR comparison: uses dynamic tolerance
                            ncr_icb = round(cons['ncr'], 2)
                            ncr_fc = round(fd['ncr'], 2)
                            if abs(ncr_icb - ncr_fc) > gap_tolerance:
                                details_ncr.append(f"{label} : ICB = {ncr_icb} / Forecast = {ncr_fc}")

                            # Cost comparison
                            ba_icb = round(cons['total_bill_amount'], 2)
                            ba_fc = round(fd['bill_amount'], 2)
                            if abs(ba_icb - ba_fc) > gap_tolerance:
                                details_ba.append((abs(ba_icb - ba_fc), label, ba_icb, ba_fc))

                        parts = []

                        if details_wd:
                            parts.append(f"Gap due to working days:\n{_limit_list(details_wd)}")
                        if details_ncr:
                            parts.append(f"Gap due to rate (NCR):\n{_limit_list(details_ncr)}")
                        if details_ba:
                            details_ba.sort(key=lambda x: x[0], reverse=True)
                            entries = [f"{x[1]} : ICB = {x[2]} / Forecast = {x[3]}" for x in details_ba]
                            parts.append(f"Cost issue:\n{_limit_list(entries)}")

                        if not parts:
                            parts.append("Cost issue (requires manual verification)")

                        comment = "\n\n".join(parts)

                    row_data = [
                        "", # 1. EM_Lead
                        cc['project_name'] or "", # 2. Project Name
                        cc['cc_project_number'], # 3. CC Project Number
                        cc['totals']['working_days'], # 4. Working Days
                        cc['totals']['bill_amount'], # 5. ICB Transfer Price
                        forecast_cell_value, # 6. Forecast Transfer Price
                        f"=E{current_row}-F{current_row}", # 7. Gap (signed, no abs)
                        cc['totals']['raw_cost'], # 8. Raw Cost
                        "", # 9. Issue Type
                    ]
                    if include_comments:
                        row_data.append(comment) # 10. Generated comment
                    ws.append(row_data)

                    # Alignment
                    for col in range(1, num_cols + 1):
                        cell = ws.cell(row=current_row, column=col)
                        if col in [4, 5, 6, 7, 8]:
                            cell.alignment = right_align
                        elif include_comments and col == num_cols:
                            # Multi-line comment: wrap_text + vertical top
                            cell.alignment = Alignment(
                                wrap_text=True,
                                vertical='top',
                                horizontal='left',
                            )
                            # Calculer la hauteur de ligne necessaire
                            if comment:
                                num_lines = comment.count("\n") + 1
                                needed_height = max(15, num_lines * 15)
                                current_height = ws.row_dimensions[current_row].height or 15
                                ws.row_dimensions[current_row].height = max(
                                    current_height, needed_height
                                )

                    current_row += 1

                # Define the Excel table
                if current_row > 2:
                    table_ref = f"A1:{get_column_letter(num_cols)}{current_row - 1}"
                    tab = Table(displayName=f"Table_{month.replace('-', '_')}", ref=table_ref)
                    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                           showLastColumn=False, showRowStripes=True, showColumnStripes=False)
                    tab.tableStyleInfo = style
                    ws.add_table(tab)

                # Column width adjustments
                ws.column_dimensions['A'].width = 25 # EM_Lead
                ws.column_dimensions['B'].width = 40 # Project Name
                ws.column_dimensions['C'].width = 25 # CC Project Number
                ws.column_dimensions['D'].width = 15 # Working Days
                ws.column_dimensions['E'].width = 20 # ICB Transfer Price
                ws.column_dimensions['F'].width = 22 # Forecast Transfer Price
                ws.column_dimensions['G'].width = 15 # Gap
                ws.column_dimensions['H'].width = 15 # Raw Cost
                ws.column_dimensions['I'].width = 25 # Issue Type
                if include_comments:
                    ws.column_dimensions['J'].width = 60 # Generated comment

            # =================================================================
            # FEUILLE KPI DASHBOARD
            # =================================================================
            self._build_kpi_sheet(wb, months_in_range, start_date, end_date, gap_tolerance)

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    def _build_kpi_sheet(self, wb, months_in_range, start_date, end_date, gap_tolerance):
        """Construit la feuille KPI Dashboard dans le workbook d'export.
        
        Sections :
        1. Synthese mensuelle (Cost, Forecast, Gap, WD, Raw Cost, Nb projets, Nb consultants)
        2. Bill Amount Cumule par mois (progression mois apres mois)
        3. Synthese par projet (toutes periodes confondues)
        4. Top consultants par Cost
        5. Detail consultant par mois
        """
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
        from openpyxl.utils import get_column_letter

        ws = wb.create_sheet(title="KPI Dashboard")

        # ---- Styles reutilisables ----
        title_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        section_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        normal_font = Font(name="Calibri", size=10)
        number_font = Font(name="Calibri", size=10)

        dark_fill = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
        section_fill = PatternFill(start_color="2E4057", end_color="2E4057", fill_type="solid")
        header_fill = PatternFill(start_color="3A5A8C", end_color="3A5A8C", fill_type="solid")
        accent_fill = PatternFill(start_color="1A73E8", end_color="1A73E8", fill_type="solid")
        light_row_fill = PatternFill(start_color="EBF0F7", end_color="EBF0F7", fill_type="solid")
        cumul_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
        negative_fill = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")
        positive_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")

        center_align = Alignment(horizontal="center", vertical="center")
        right_align = Alignment(horizontal="right", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")
        wrap_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        thin_border = Border(
            left=Side(style="thin", color="B0BEC5"),
            right=Side(style="thin", color="B0BEC5"),
            top=Side(style="thin", color="B0BEC5"),
            bottom=Side(style="thin", color="B0BEC5"),
        )

        num_fmt_money = '#,##0.00'
        num_fmt_days = '#,##0.00'
        num_fmt_int = '#,##0'
        num_fmt_pct = '0.0%'

        def _style_header_row(ws, row, col_start, col_end, fill=None):
            f = fill or header_fill
            for c in range(col_start, col_end + 1):
                cell = ws.cell(row=row, column=c)
                cell.font = header_font
                cell.fill = f
                cell.alignment = center_align
                cell.border = thin_border

        def _style_data_cell(ws, row, col, fmt=None, align=None, fill=None):
            cell = ws.cell(row=row, column=col)
            cell.font = number_font
            cell.border = thin_border
            if fmt:
                cell.number_format = fmt
            if align:
                cell.alignment = align
            else:
                cell.alignment = right_align
            if fill:
                cell.fill = fill
            return cell

        def _write_section_title(ws, row, title, col_span=8):
            cell = ws.cell(row=row, column=1, value=title)
            cell.font = section_font
            cell.fill = section_fill
            cell.alignment = Alignment(horizontal="left", vertical="center")
            for c in range(1, col_span + 1):
                ws.cell(row=row, column=c).fill = section_fill
                ws.cell(row=row, column=c).border = thin_border

        # ---- Collecter toutes les donnees par mois ----
        monthly_stats = {}  # month -> {cost, forecast, gap, wd, raw_cost, nb_proj, nb_cons}
        project_stats = {}  # cc_proj -> {name, cost, forecast, gap, wd, raw_cost, nb_cons, avg_ncr}
        consultant_monthly = {}  # (emp_id, name) -> {month -> {cost, wd, ncr, project}}

        for month in months_in_range:
            data = self.hierarchy_data(month_filter=month, start_date=start_date, end_date=end_date)
            m_cost = 0.0
            m_forecast = 0.0
            m_wd = 0.0
            m_raw_cost = 0.0
            m_nb_proj = 0
            m_nb_cons = 0

            for cc in data:
                m_nb_proj += 1
                norm_proj = ExpenditureItem._normalize_key(cc['cc_project_number'])
                proj_key = cc['cc_project_number']

                # Forecast pour ce projet/mois
                fc_total = 0.0
                has_fc = False
                for fk, fd in self.forecast_data.items():
                    if fk[0] == norm_proj and fk[2] == month:
                        if abs(fd.get("working_days", 0)) > 1e-6 or abs(fd.get("bill_amount", 0)) > 1e-6:
                            has_fc = True
                            fc_total += fd["bill_amount"]

                m_cost += cc['totals']['bill_amount']
                m_forecast += fc_total if has_fc else 0
                m_wd += cc['totals']['working_days']
                m_raw_cost += cc['totals']['raw_cost']
                m_nb_cons += len(cc['consultants'])

                # Agreger par projet
                if proj_key not in project_stats:
                    project_stats[proj_key] = {
                        "name": cc['project_name'] or "",
                        "cost": 0.0, "forecast": 0.0, "wd": 0.0,
                        "raw_cost": 0.0, "nb_cons": set(), "ncr_sum": 0.0,
                        "ncr_count": 0, "months_active": []
                    }
                ps = project_stats[proj_key]
                ps["cost"] += cc['totals']['bill_amount']
                ps["forecast"] += fc_total if has_fc else 0
                ps["wd"] += cc['totals']['working_days']
                ps["raw_cost"] += cc['totals']['raw_cost']
                ps["months_active"].append(month)

                for cons in cc['consultants']:
                    ps["nb_cons"].add(cons['employee_id'])
                    if cons['ncr'] and cons['ncr'] > 0:
                        ps["ncr_sum"] += cons['ncr']
                        ps["ncr_count"] += 1

                    # Detail consultant mensuel
                    cons_key = (cons['employee_id'], cons['name'] or "")
                    if cons_key not in consultant_monthly:
                        consultant_monthly[cons_key] = {}
                    if month not in consultant_monthly[cons_key]:
                        consultant_monthly[cons_key][month] = {
                            "cost": 0.0, "wd": 0.0, "ncr": 0.0,
                            "project": cc['project_name'] or proj_key
                        }
                    cm = consultant_monthly[cons_key][month]
                    cm["cost"] += cons['total_bill_amount']
                    cm["wd"] += cons['working_days']
                    if cons['ncr'] and cons['ncr'] > 0:
                        cm["ncr"] = cons['ncr']

            monthly_stats[month] = {
                "cost": round(m_cost, 2),
                "forecast": round(m_forecast, 2),
                "gap": round(m_cost - m_forecast, 2),
                "wd": round(m_wd, 2),
                "raw_cost": round(m_raw_cost, 2),
                "nb_proj": m_nb_proj,
                "nb_cons": m_nb_cons,
            }

        # ================================================================
        # SECTION 0 : TITRE PRINCIPAL
        # ================================================================
        row = 1
        title_cell = ws.cell(row=row, column=1, value="KPI DASHBOARD — ICB Analysis")
        title_cell.font = title_font
        title_cell.fill = dark_fill
        for c in range(1, 10):
            ws.cell(row=row, column=c).fill = dark_fill
        row += 1
        period_label = ""
        if start_date and end_date:
            period_label = f"Periode : {start_date} a {end_date}"
        elif months_in_range:
            period_label = f"Mois : {months_in_range[0]} a {months_in_range[-1]}"
        sub_cell = ws.cell(row=row, column=1, value=period_label)
        sub_cell.font = Font(name="Calibri", size=10, italic=True, color="8899AA")
        sub_cell.fill = dark_fill
        for c in range(1, 10):
            ws.cell(row=row, column=c).fill = dark_fill
        row += 2

        # ================================================================
        # SECTION 1 : SYNTHESE MENSUELLE
        # ================================================================
        _write_section_title(ws, row, "1. SYNTHESE MENSUELLE", 8)
        row += 1
        headers_s1 = ["Mois", "ICB Transfer Price", "Forecast Transfer Price", "Gap (ICB-Forecast)",
                       "Working Days", "Raw Cost", "Nb Projets", "Nb Consultants"]
        for c, h in enumerate(headers_s1, 1):
            ws.cell(row=row, column=c, value=h)
        _style_header_row(ws, row, 1, len(headers_s1))
        row += 1

        for m in months_in_range:
            ms = monthly_stats[m]
            ws.cell(row=row, column=1, value=m)
            _style_data_cell(ws, row, 1, align=center_align)
            ws.cell(row=row, column=2, value=ms["cost"])
            _style_data_cell(ws, row, 2, fmt=num_fmt_money)
            ws.cell(row=row, column=3, value=ms["forecast"])
            _style_data_cell(ws, row, 3, fmt=num_fmt_money)
            ws.cell(row=row, column=4, value=ms["gap"])
            gap_fill = negative_fill if ms["gap"] < 0 else (positive_fill if ms["gap"] > gap_tolerance else None)
            _style_data_cell(ws, row, 4, fmt=num_fmt_money, fill=gap_fill)
            ws.cell(row=row, column=5, value=ms["wd"])
            _style_data_cell(ws, row, 5, fmt=num_fmt_days)
            ws.cell(row=row, column=6, value=ms["raw_cost"])
            _style_data_cell(ws, row, 6, fmt=num_fmt_money)
            ws.cell(row=row, column=7, value=ms["nb_proj"])
            _style_data_cell(ws, row, 7, fmt=num_fmt_int, align=center_align)
            ws.cell(row=row, column=8, value=ms["nb_cons"])
            _style_data_cell(ws, row, 8, fmt=num_fmt_int, align=center_align)
            row += 1

        # Ligne TOTAL
        ws.cell(row=row, column=1, value="TOTAL")
        _style_data_cell(ws, row, 1, align=center_align,
                         fill=accent_fill)
        ws.cell(row=row, column=1).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        for c in range(2, 9):
            ws.cell(row=row, column=c).fill = accent_fill
            ws.cell(row=row, column=c).font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            ws.cell(row=row, column=c).border = thin_border
            ws.cell(row=row, column=c).alignment = right_align
        total_cost = sum(ms["cost"] for ms in monthly_stats.values())
        total_fc = sum(ms["forecast"] for ms in monthly_stats.values())
        total_gap = round(total_cost - total_fc, 2)
        total_wd = sum(ms["wd"] for ms in monthly_stats.values())
        total_rc = sum(ms["raw_cost"] for ms in monthly_stats.values())
        ws.cell(row=row, column=2, value=round(total_cost, 2))
        ws.cell(row=row, column=2).number_format = num_fmt_money
        ws.cell(row=row, column=3, value=round(total_fc, 2))
        ws.cell(row=row, column=3).number_format = num_fmt_money
        ws.cell(row=row, column=4, value=total_gap)
        ws.cell(row=row, column=4).number_format = num_fmt_money
        ws.cell(row=row, column=5, value=round(total_wd, 2))
        ws.cell(row=row, column=5).number_format = num_fmt_days
        ws.cell(row=row, column=6, value=round(total_rc, 2))
        ws.cell(row=row, column=6).number_format = num_fmt_money
        ws.cell(row=row, column=7, value="")
        ws.cell(row=row, column=7).alignment = center_align
        ws.cell(row=row, column=8, value="")
        ws.cell(row=row, column=8).alignment = center_align
        row += 2

        # ================================================================
        # SECTION 2 : TRANSFER PRICE CUMULE PAR MOIS
        # ================================================================
        _write_section_title(ws, row, "2. TRANSFER PRICE CUMULE PAR MOIS (Progression)", 8)
        row += 1
        headers_s2 = ["Mois", "ICB TP du Mois", "ICB TP Cumule",
                       "Forecast TP du Mois", "Forecast TP Cumule",
                       "Gap Cumule", "% Progression TP"]
        for c, h in enumerate(headers_s2, 1):
            ws.cell(row=row, column=c, value=h)
        _style_header_row(ws, row, 1, len(headers_s2))
        row += 1

        cumul_cost = 0.0
        cumul_fc = 0.0
        prev_cumul_cost = 0.0
        for m in months_in_range:
            ms = monthly_stats[m]
            cumul_cost += ms["cost"]
            cumul_fc += ms["forecast"]
            pct_progression = ((cumul_cost - prev_cumul_cost) / prev_cumul_cost) if prev_cumul_cost > 0 else None

            ws.cell(row=row, column=1, value=m)
            _style_data_cell(ws, row, 1, align=center_align)
            ws.cell(row=row, column=2, value=ms["cost"])
            _style_data_cell(ws, row, 2, fmt=num_fmt_money)
            ws.cell(row=row, column=3, value=round(cumul_cost, 2))
            _style_data_cell(ws, row, 3, fmt=num_fmt_money, fill=cumul_fill)
            ws.cell(row=row, column=3).font = Font(name="Calibri", size=10, bold=True, color="2E7D32")
            ws.cell(row=row, column=4, value=ms["forecast"])
            _style_data_cell(ws, row, 4, fmt=num_fmt_money)
            ws.cell(row=row, column=5, value=round(cumul_fc, 2))
            _style_data_cell(ws, row, 5, fmt=num_fmt_money, fill=cumul_fill)
            ws.cell(row=row, column=5).font = Font(name="Calibri", size=10, bold=True, color="1565C0")
            ws.cell(row=row, column=6, value=round(cumul_cost - cumul_fc, 2))
            gap_cumul_fill = negative_fill if (cumul_cost - cumul_fc) < 0 else positive_fill
            _style_data_cell(ws, row, 6, fmt=num_fmt_money, fill=gap_cumul_fill)
            if pct_progression is not None:
                ws.cell(row=row, column=7, value=round(pct_progression, 4))
                _style_data_cell(ws, row, 7, fmt=num_fmt_pct, align=center_align)
            else:
                ws.cell(row=row, column=7, value="—")
                _style_data_cell(ws, row, 7, align=center_align)

            prev_cumul_cost = cumul_cost
            row += 1

        row += 1

        # ================================================================
        # SECTION 3 : SYNTHESE PAR PROJET
        # ================================================================
        _write_section_title(ws, row, "3. SYNTHESE PAR PROJET (toutes periodes)", 9)
        row += 1
        headers_s3 = ["CC Project Number", "Project Name", "ICB TP Total",
                       "Forecast TP Total", "Gap", "Gap %",
                       "Working Days", "NCR Moyen", "Nb Consultants"]
        for c, h in enumerate(headers_s3, 1):
            ws.cell(row=row, column=c, value=h)
        _style_header_row(ws, row, 1, len(headers_s3))
        row += 1

        # Trier par cost decroissant
        sorted_projects = sorted(project_stats.items(),
                                 key=lambda x: x[1]["cost"], reverse=True)
        for proj_key, ps in sorted_projects:
            gap = round(ps["cost"] - ps["forecast"], 2)
            gap_pct = (gap / ps["forecast"]) if ps["forecast"] != 0 else None
            avg_ncr = round(ps["ncr_sum"] / ps["ncr_count"], 2) if ps["ncr_count"] > 0 else 0

            ws.cell(row=row, column=1, value=proj_key)
            _style_data_cell(ws, row, 1, align=left_align)
            ws.cell(row=row, column=2, value=ps["name"])
            _style_data_cell(ws, row, 2, align=left_align)
            ws.cell(row=row, column=3, value=round(ps["cost"], 2))
            _style_data_cell(ws, row, 3, fmt=num_fmt_money)
            ws.cell(row=row, column=4, value=round(ps["forecast"], 2))
            _style_data_cell(ws, row, 4, fmt=num_fmt_money)
            ws.cell(row=row, column=5, value=gap)
            g_fill = negative_fill if gap < 0 else (positive_fill if gap > gap_tolerance else None)
            _style_data_cell(ws, row, 5, fmt=num_fmt_money, fill=g_fill)
            if gap_pct is not None:
                ws.cell(row=row, column=6, value=round(gap_pct, 4))
                _style_data_cell(ws, row, 6, fmt=num_fmt_pct, align=center_align)
            else:
                ws.cell(row=row, column=6, value="N/A")
                _style_data_cell(ws, row, 6, align=center_align)
            ws.cell(row=row, column=7, value=round(ps["wd"], 2))
            _style_data_cell(ws, row, 7, fmt=num_fmt_days)
            ws.cell(row=row, column=8, value=avg_ncr)
            _style_data_cell(ws, row, 8, fmt=num_fmt_money)
            ws.cell(row=row, column=9, value=len(ps["nb_cons"]))
            _style_data_cell(ws, row, 9, fmt=num_fmt_int, align=center_align)

            # Alternance de couleur
            if (row % 2) == 0:
                for c in range(1, 10):
                    if ws.cell(row=row, column=c).fill == PatternFill():
                        ws.cell(row=row, column=c).fill = light_row_fill

            row += 1

        row += 1

        # ================================================================
        # SECTION 4 : TOP CONSULTANTS PAR COST
        # ================================================================
        _write_section_title(ws, row, "4. TOP CONSULTANTS PAR TRANSFER PRICE (toutes periodes)", 8)
        row += 1
        headers_s4 = ["Employee ID", "Nom", "Projet Principal",
                       "ICB TP Total", "Working Days Total",
                       "NCR", "Nb Mois Actifs", "TP Moyen / Mois"]
        for c, h in enumerate(headers_s4, 1):
            ws.cell(row=row, column=c, value=h)
        _style_header_row(ws, row, 1, len(headers_s4))
        row += 1

        # Agreger par consultant
        cons_totals = []
        for (emp_id, name), months_data in consultant_monthly.items():
            total_cost = sum(md["cost"] for md in months_data.values())
            total_wd = sum(md["wd"] for md in months_data.values())
            ncr_vals = [md["ncr"] for md in months_data.values() if md["ncr"] > 0]
            avg_ncr = round(sum(ncr_vals) / len(ncr_vals), 2) if ncr_vals else 0
            # Projet le plus frequent
            proj_counts = {}
            for md in months_data.values():
                proj_counts[md["project"]] = proj_counts.get(md["project"], 0) + 1
            main_proj = max(proj_counts, key=proj_counts.get) if proj_counts else ""
            nb_months = len(months_data)
            avg_cost_month = round(total_cost / nb_months, 2) if nb_months > 0 else 0

            cons_totals.append({
                "emp_id": emp_id, "name": name, "project": main_proj,
                "cost": round(total_cost, 2), "wd": round(total_wd, 2),
                "ncr": avg_ncr, "nb_months": nb_months,
                "avg_cost_month": avg_cost_month,
            })

        # Trier par cost decroissant, limiter a 50
        cons_totals.sort(key=lambda x: x["cost"], reverse=True)
        for ct in cons_totals[:50]:
            ws.cell(row=row, column=1, value=ct["emp_id"])
            _style_data_cell(ws, row, 1, align=left_align)
            ws.cell(row=row, column=2, value=ct["name"])
            _style_data_cell(ws, row, 2, align=left_align)
            ws.cell(row=row, column=3, value=ct["project"])
            _style_data_cell(ws, row, 3, align=left_align)
            ws.cell(row=row, column=4, value=ct["cost"])
            _style_data_cell(ws, row, 4, fmt=num_fmt_money)
            ws.cell(row=row, column=5, value=ct["wd"])
            _style_data_cell(ws, row, 5, fmt=num_fmt_days)
            ws.cell(row=row, column=6, value=ct["ncr"])
            _style_data_cell(ws, row, 6, fmt=num_fmt_money)
            ws.cell(row=row, column=7, value=ct["nb_months"])
            _style_data_cell(ws, row, 7, fmt=num_fmt_int, align=center_align)
            ws.cell(row=row, column=8, value=ct["avg_cost_month"])
            _style_data_cell(ws, row, 8, fmt=num_fmt_money)

            if (row % 2) == 0:
                for c in range(1, 9):
                    if ws.cell(row=row, column=c).fill == PatternFill():
                        ws.cell(row=row, column=c).fill = light_row_fill
            row += 1

        row += 1

        # ================================================================
        # SECTION 5 : DETAIL CONSULTANT PAR MOIS
        # ================================================================
        _write_section_title(ws, row,
                             "5. DETAIL CONSULTANT PAR MOIS (TP par consultant par mois)",
                             3 + len(months_in_range))
        row += 1

        # Headers : Employee ID | Nom | Projet | Mois1 | Mois2 | ... | TOTAL
        headers_s5 = ["Employee ID", "Nom", "Projet"] + months_in_range + ["TOTAL"]
        for c, h in enumerate(headers_s5, 1):
            ws.cell(row=row, column=c, value=h)
        _style_header_row(ws, row, 1, len(headers_s5))
        row += 1

        # Trier par nom
        sorted_cons = sorted(consultant_monthly.items(),
                             key=lambda x: x[0][1].lower() if x[0][1] else x[0][0])
        for (emp_id, name), months_data in sorted_cons:
            proj_counts = {}
            for md in months_data.values():
                proj_counts[md["project"]] = proj_counts.get(md["project"], 0) + 1
            main_proj = max(proj_counts, key=proj_counts.get) if proj_counts else ""

            ws.cell(row=row, column=1, value=emp_id)
            _style_data_cell(ws, row, 1, align=left_align)
            ws.cell(row=row, column=2, value=name)
            _style_data_cell(ws, row, 2, align=left_align)
            ws.cell(row=row, column=3, value=main_proj)
            _style_data_cell(ws, row, 3, align=left_align)

            row_total = 0.0
            for mi, m in enumerate(months_in_range):
                col = 4 + mi
                if m in months_data:
                    val = round(months_data[m]["cost"], 2)
                    ws.cell(row=row, column=col, value=val)
                    row_total += val
                else:
                    ws.cell(row=row, column=col, value=0)
                _style_data_cell(ws, row, col, fmt=num_fmt_money)

            # Colonne TOTAL
            total_col = 4 + len(months_in_range)
            ws.cell(row=row, column=total_col, value=round(row_total, 2))
            _style_data_cell(ws, row, total_col, fmt=num_fmt_money,
                             fill=cumul_fill)
            ws.cell(row=row, column=total_col).font = Font(
                name="Calibri", size=10, bold=True, color="2E7D32")

            if (row % 2) == 0:
                for c in range(1, total_col + 1):
                    if ws.cell(row=row, column=c).fill == PatternFill():
                        ws.cell(row=row, column=c).fill = light_row_fill
            row += 1

        # ---- Largeurs de colonnes ----
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['B'].width = 35
        ws.column_dimensions['C'].width = 22
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 18
        ws.column_dimensions['F'].width = 18
        ws.column_dimensions['G'].width = 18
        ws.column_dimensions['H'].width = 18
        ws.column_dimensions['I'].width = 18

        # Figer la premiere ligne
        ws.freeze_panes = "A2"


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


@app.route("/api/reset", methods=["POST"])
def reset_data():
    """Remet l'application a zero : vide le pipeline et supprime les fichiers temporaires."""
    global pipeline
    pipeline = None

    # Supprimer les fichiers uploades
    upload_dir = os.path.join(os.path.dirname(__file__), "uploads")
    for fname in ("current_icb.xlsx", "current_forecast.xlsx"):
        fpath = os.path.join(upload_dir, fname)
        if os.path.exists(fpath):
            try:
                os.remove(fpath)
            except OSError:
                pass

    return jsonify({"message": "Donnees dechargees avec succes."})


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
    
    gap_tolerance_str = request.args.get("gap_tolerance", "2")
    try:
        gap_tolerance = float(gap_tolerance_str)
    except ValueError:
        gap_tolerance = 2.0
    
    # Whether to include the "Generated comment" column
    include_comments = request.args.get("include_comments", "1") in ("1", "true", "yes")
    
    try:
        buffer = pipeline.export_to_excel(month_filter=month, start_date=start_date, end_date=end_date, gap_tolerance=gap_tolerance, include_comments=include_comments)
        
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

    