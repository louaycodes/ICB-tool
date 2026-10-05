"""
ICB Application - Capgemini
Flask backend: loads the Excel file, filters, groups, and exposes data
via JSON endpoints for the web interface.
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
    """ Get the absolute path to the resource, works in dev and with PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Flask initialization with paths managed by PyInstaller
app = Flask(
    __name__,
    template_folder=get_resource_path('templates'),
    static_folder=get_resource_path('static')
)

# If you access the uploads folder in your code, also use:
UPLOAD_FOLDER = get_resource_path('uploads')

# --- Business parameters ---
EXPENDITURE_TYPES = [
    "RC_Time Std",
    "RC_Overtime Std Cost",
    "RC_Overtime Std Cost x2",
    "RC_Overtime Std Cost x1.25",
    "RC_Overtime Std Cost x1.5",
]
UOM_FILTER = "HOURS"
NCR_CORRELATION_TARGET = 0.75
NCR_CORRELATION_TOLERANCE = 0.05  # valid if abs(correlation) in [0.70, 0.80]

# Tolerance thresholds are now sent dynamically from the frontend.

# --- Flask app ---
app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB max for upload

# In-memory storage for the pipeline (single session)
pipeline = None


# ---------------------------------------------------------------------------
# Fast loading of large Excel files
# ---------------------------------------------------------------------------
def fast_load_excel(path, sheet_name=0, dynamic_header=False):
    """Load an Excel file into a DataFrame.

    If *dynamic_header* is True, the function scans the first 20 rows
    to find the one containing expected headers (same logic as the
    Forecast loading). Otherwise, row 1 is used as the header.

    Returns a pandas DataFrame.
    """
    wb = load_workbook(path, read_only=True, data_only=True)

    # --- Sheet selection ---
    if isinstance(sheet_name, str):
        if sheet_name not in wb.sheetnames:
            available = wb.sheetnames
            wb.close()
            raise ValueError(
                f"Sheet '{sheet_name}' not found in the workbook. "
                f"Available sheets: {available}"
            )
        ws = wb[sheet_name]
    else:
        ws = wb.worksheets[sheet_name]

    SCAN_MAX_COL = 200

    if dynamic_header:
        # --- Dynamic header scan ---
        # Expected markers to identify the ICB header row
        ICB_HEADER_MARKERS = {
            'employee number', 'cc project number', 'project name',
            'expenditure type', 'quantity',
        }

        header_row_idx = None   # 1-indexed
        header_cells = {}       # col_0idx -> str value
        min_data_col = None     # 0-indexed: first non-empty column

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
                "Unable to find the header row in the first 20 "
                "rows of the sheet. Searched markers: "
                + ", ".join(sorted(ICB_HEADER_MARKERS))
            )

        # Determine the first and last columns containing a header
        col_indices = sorted(header_cells.keys())
        min_data_col = col_indices[0]       # 0-indexed
        max_data_col = col_indices[-1]       # 0-indexed
        ncols = max_data_col - min_data_col + 1

        # Build the header (may contain None for gaps)
        header = []
        for ci in range(min_data_col, max_data_col + 1):
            header.append(header_cells.get(ci))

        print(f"[ICB] Headers detected at row {header_row_idx}, "
              f"columns {min_data_col}-{max_data_col} ({ncols} columns)")

        # Read data starting from the next row
        rows_iter = ws.iter_rows(
            min_row=header_row_idx + 1,
            min_col=min_data_col + 1,   # 1-indexed for openpyxl
            max_col=max_data_col + 1,   # 1-indexed
            values_only=True,
        )
        data = [r for r in rows_iter if any(v is not None for v in r)]

    else:
        # --- Classic mode: headers on row 1, data in col A ---
        header_row = next(ws.iter_rows(min_row=1, max_row=1, min_col=1,
                                        max_col=SCAN_MAX_COL, values_only=True))
        ncols = max((i for i, v in enumerate(header_row) if v is not None),
                    default=-1) + 1
        if ncols == 0:
            wb.close()
            raise ValueError("Empty sheet or header not found.")
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
# Business model
# ---------------------------------------------------------------------------
class ExpenditureItem:
    """An expenditure/billing line item."""

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
        # Correlation column for anomaly detection
        raw_corr = _clean(row.get("FULL NCR to Direct NCR Correlation"))
        self.ncr_correlation = self._safe_float(raw_corr)

    @property
    def month_key(self):
        return self.date.strftime("%Y-%m") if self.date else "Unknown"

    @property
    def has_anomaly(self):
        """True if abs(ncr_correlation) is outside [0.70, 0.80]."""
        if self.ncr_correlation is None:
            return False
        val = abs(self.ncr_correlation)
        return val < (NCR_CORRELATION_TARGET - NCR_CORRELATION_TOLERANCE) or \
               val > (NCR_CORRELATION_TARGET + NCR_CORRELATION_TOLERANCE)

    @staticmethod
    def _safe_float(value):
        """Safely converts a value to float.
        Returns None if the value is None or not convertible.
        Handles the case where Excel returns numbers as strings."""
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        # Attempt conversion from a string
        try:
            f = float(value)
            if math.isnan(f) or math.isinf(f):
                return None
            return f
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _normalize_key(value):
        """Forces all project/CC-project keys to clean text (avoids
        str/int/None mixing that breaks sorted() downstream)."""
        if value is None:
            return "N/A"
        if isinstance(value, float) and value.is_integer():
            return str(int(value))   # 123.0 -> "123", avoids duplicates with "123"
        return str(value).strip()

    def matches_filter(self):
        return self.expenditure_type in EXPENDITURE_TYPES and self.uom == UOM_FILTER


# ---------------------------------------------------------------------------
# Processing pipeline
# ---------------------------------------------------------------------------
class ICBPipeline:
    def __init__(self):
        self.items = []
        self.filtered_items = []
        self.available_months = []
        self.forecast_data = {}  # (cc_proj, emp, month_key) -> {"ncr": float, "working_days": float, "bill_amount": float}
        self.forecast_names = {}  # (norm_proj, norm_ggid) -> employee_name (from forecast file)
        self.project_em_map = {}  # norm_proj -> em_name (from EDM sheet)

    def load(self, path):
        # Load the "ICB_DB" sheet with dynamic header detection
        df_raw = fast_load_excel(path, sheet_name="DB_ICB", dynamic_header=True)

        # --- Dynamic detection of the "Project Transfer Price" column ---
        # Accepts 3 ordered candidate name variants:
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
                    print(f"[INFO] Alternative column '{ptp_col}' detected "
                          f"(used instead of '{candidate_names[0]}')")
                break

        if ptp_col is None:
            names_str = ", ".join(f"'{name}'" for name in candidate_names)
            print(f"[ERROR] None of the expected columns ({names_str}) "
                  "were found in the ICB file.")
            print(f"  Available columns: {list(df_raw.columns)}")
        elif str(ptp_col).strip() != "Project Transfer Price":
            # Rename to the canonical name expected by ExpenditureItem
            df_raw = df_raw.rename(columns={ptp_col: "Project Transfer Price"})
            print(f"[INFO] Column '{ptp_col}' renamed to 'Project Transfer Price'")

        self.items = [ExpenditureItem(row) for row in df_raw.to_dict("records")]

    def load_forecast(self, path):
        """Parse a Forecast file and populate self.forecast_data.

        Dynamic structure detection:
        - Header row found by scanning the first 20 rows
        - Identity columns (Project Num, GGID, …) found by name
        - Monthly blocks located via month names (FR/EN) and
          sub-columns (Grade, F.RC, Vaca, Prod, Cost, FTE)
        - Total / TR Q* blocks excluded automatically
        """
        import openpyxl
        from datetime import datetime

        # ---- Local helpers ----
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

        SCAN_MAX_COL = 200  # do not scan phantom Excel columns

        # Expected identity markers in the header row
        IDENTITY_MARKERS = {'project num', 'project name', 'icb id',
                            'employee name', 'grade'}

        # Mapping month names FR + EN -> month number (1-12)
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

        # Expected names in a monthly block (canonical order)
        EXPECTED_SUB_ORDER = ['grade', 'f.rc', 'vaca', 'prod', 'cost', 'fte']
        SUB_COL_SET = set(EXPECTED_SUB_ORDER)

        # Parasite values: if Project Num == one of these words, it's a
        # misidentified line (duplicate header, etc.), not data.
        PARASITE_VALUES = (
            IDENTITY_MARKERS
            | set(MONTH_NAMES_MAP.keys())
            | SUB_COL_SET
            | {'total'}
        )

        print("\n" + "=" * 70)
        print("[FORECAST IMPORT] Starting Forecast file loading")
        print("=" * 70)

        # =================================================================
        # Open in normal mode for structure detection
        # (merged_cells, dimensions, …). Forecast files = small,
        # so no performance issue.
        # =================================================================
        wb = openpyxl.load_workbook(path, data_only=True)

        print(f"\n  Sheets present in the workbook: {wb.sheetnames}")

        # =================================================================
        # 1. DETECTING THE SHEET AND MAIN HEADER ROW
        # =================================================================
        print("\n--- 1. DETECTING THE SHEET AND MAIN HEADER ROW ---")

        ws = None
        header_row_idx = None   # 1-indexed
        header_cells = {}       # col_0idx -> str value (stripped)

        for sname in wb.sheetnames:
            # Explicitly exclude the EDM sheet from the main sheet search
            if sname.strip().lower() in ("edm", "project-em"):
                continue

            candidate_ws = wb[sname]
            for r in range(1, min(candidate_ws.max_row + 1, 21)):
                row_map = {}
                hits = 0
                for cell in next(candidate_ws.iter_rows(min_row=r, max_row=r,
                                              max_col=SCAN_MAX_COL,
                                              values_only=False)):
                    if cell.value is not None and isinstance(cell.value, str):
                        cleaned = cell.value.strip()
                        if cleaned:
                            row_map[cell.column - 1] = cleaned
                            if cleaned.lower() in IDENTITY_MARKERS:
                                hits += 1
                if hits >= 3:
                    ws = candidate_ws
                    header_row_idx = r
                    header_cells = row_map
                    break
            if ws is not None:
                break

        if ws is None or header_row_idx is None:
            print("  [ERROR] Unable to find the main sheet with its headers in the first 20 rows.")
            print("  Searched markers: "
                  + ", ".join(sorted(IDENTITY_MARKERS)))
            print("  Import cancelled (no data loaded).")
            wb.close()
            return

        main_sheet_title = ws.title
        print(f"  Main sheet detected: '{main_sheet_title}'")
        print(f"  Dimensions          : {ws.dimensions}")
        print(f"  Header row found    : row {header_row_idx}")
        print(f"  Columns with text in this row: {len(header_cells)}")

        # =================================================================
        # 2. DYNAMIC DETECTION OF IDENTITY COLUMNS
        # =================================================================
        print("\n--- 2. DETECTING IDENTITY COLUMNS ---")

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

        # --- Special detection for ICB ID (may be in a hidden column) ---
        if 'icb_id' not in identity_cols:
            print("  [INFO] 'ICB ID' not found in visible columns, "
                  "searching in hidden columns...")
            for cell in next(ws.iter_rows(min_row=header_row_idx,
                                          max_row=header_row_idx,
                                          max_col=SCAN_MAX_COL,
                                          values_only=False)):
                if cell.value is not None and isinstance(cell.value, str):
                    if cell.value.strip().lower() == 'icb id':
                        c0 = cell.column - 1
                        col_letter = _col_letter(c0)
                        # Check if the column is hidden
                        is_hidden = False
                        try:
                            dim = ws.column_dimensions.get(col_letter)
                            if dim and dim.hidden:
                                is_hidden = True
                        except Exception:
                            pass
                        identity_cols['icb_id'] = c0
                        print(f"  [OK] 'ICB ID' found in column {col_letter} "
                              f"(0-idx {c0})"
                              f"{' [HIDDEN COLUMN]' if is_hidden else ''}")
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
                print(f"  [OK] '{label}' -> column {_col_letter(idx)} "
                      f"(0-idx {idx})")
            else:
                print(f"  [WARNING] '{label}' not found in the header "
                      f"row")

        if col_proj is None or col_icb_id is None:
            print("  [ERROR] Required columns missing "
                  "(Project Num and/or ICB ID). Import cancelled.")
            wb.close()
            return

        # =================================================================
        # 3. DYNAMIC DETECTION OF MONTHLY BLOCKS
        # =================================================================
        print("\n--- 3. DETECTING MONTHLY BLOCKS ---")

        # 3a. Scan the rows above the header (and the header itself)
        #     to locate month / Total / TR labels.
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

        # 3b. For each month label, identify the sub-columns
        #     (Grade, F.RC, Vaca, Prod, Cost, FTE) in the header.
        month_blocks = {}  # month_num -> {frc: c0, prod: c0, cost: c0}

        for label_col in sorted(block_labels.keys()):
            label_type, label_val = block_labels[label_col]

            if label_type == 'skip':
                print(f"  [SKIP] Block '{label_val}' at col "
                      f"{_col_letter(label_col)} (Total/TR)")
                continue

            month_num = label_val

            # Check the 6 columns starting at label_col in the header
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
                # Fallback: search each sub-column by name in
                # an expanded range (label_col .. label_col+7)
                print(f"  [WARNING] Month {month_num:02d} "
                      f"(col {_col_letter(label_col)}): non-standard "
                      f"sub-column order, searching by name")
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
                print(f"  [WARNING] Month {month_num:02d}: missing columns "
                      f"{missing} — this month will be skipped")
                continue

            month_blocks[month_num] = {
                'frc': frc_col, 'prod': prod_col, 'cost': cost_col
            }
            print(f"  [OK] Month {month_num:02d} : "
                  f"F.RC=col {_col_letter(frc_col)}, "
                  f"Prod=col {_col_letter(prod_col)}, "
                  f"Cost=col {_col_letter(cost_col)}")

        if not month_blocks:
            print("  [ERROR] No valid monthly block detected. "
                  "Import cancelled.")
            wb.close()
            return

        print(f"  => {len(month_blocks)} valid monthly block(s)")

        # =================================================================
        # 4. DETECTING PROJECT BLOCKS
        #    Search for merged ranges in the dynamically detected
        #    Project Num column, or fallback by value change.
        # =================================================================
        print("\n--- 4. DETECTING PROJECT BLOCKS ---")

        proj_col_1idx = col_proj + 1   # 1-indexed for the merged_cells API
        icb_id_col_1idx = col_icb_id + 1
        emp_col_1idx = (col_emp_name + 1) if col_emp_name is not None else None

        merged_ranges = sorted(
            [mr for mr in ws.merged_cells.ranges
             if mr.min_col == proj_col_1idx],
            key=lambda r: r.min_row,
        )

        if merged_ranges:
            print(f"  {len(merged_ranges)} merged range(s) in column "
                  f"{_col_letter(col_proj)}:")
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
                    f"    Project '{anchor}': rows {mr.min_row}-{mr.max_row} "
                    f"({n_rows} consultants), "
                    f"Total row={tot_row} "
                    f"{'[detected and excluded]' if has_total else '[NOT detected — WARNING]'}"
                )
        else:
            print(f"  [INFO] No merged range in column "
                  f"{_col_letter(col_proj)}.")
            print("  Fallback: detection by value change:")
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
                    print(f"    Project '{pv}': {len(rows)} consultant(s) "
                          f"(rows {rows[0]}-{rows[-1]})")
            else:
                print("    [WARNING] No Project Num found.")

        wb.close()

        # =================================================================
        # MAIN PARSING (read_only for performance)
        # =================================================================
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb[main_sheet_title]
            current_year = datetime.now().year
            current_proj = None

            # Diagnostic counters
            projects_found = set()
            consultants_by_project = {}
            rows_processed = 0
            rows_skipped = 0

            data_start_row = header_row_idx + 1

            for row in ws.iter_rows(min_row=data_start_row, values_only=True):
                # Safe access to dynamically detected columns
                def _safe(idx):
                    if idx is not None and idx < len(row):
                        return row[idx]
                    return None

                proj_val = _safe(col_proj)
                icb_id_val = _safe(col_icb_id)

                # Update current project
                if proj_val is not None:
                    # Guard: ignore parasite lines
                    pv_lower = str(proj_val).strip().lower()
                    if pv_lower in PARASITE_VALUES:
                        continue
                    current_proj = proj_val

                # Ignore Total lines (ICB ID empty)
                if icb_id_val is None:
                    rows_skipped += 1
                    continue

                rows_processed += 1
                norm_proj = ExpenditureItem._normalize_key(current_proj)
                norm_icb_id = str(icb_id_val).strip().replace("6427_", "")

                # Store the consultant name from the forecast
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

                    # Skip if all values are 0 or empty
                    if abs(frc) > 1e-6 or abs(prod) > 1e-6 or abs(cost) > 1e-6:
                        self.forecast_data[
                            (norm_proj, norm_icb_id, month_key)
                        ] = {
                            "ncr": frc,
                            "working_days": prod,
                            "bill_amount": cost,
                        }

            # =============================================================
            # 5. MAPPING WITH THE ICB FILE
            # =============================================================
            print("\n--- 5. MAPPING WITH THE ICB FILE ---")

            icb_projects = set()
            icb_employees_by_project = {}
            for item in self.filtered_items:
                p = ExpenditureItem._normalize_key(item.cc_project_number)
                e = str(item.employee_number).strip().replace("6427_", "")
                icb_projects.add(p)
                icb_employees_by_project.setdefault(p, set()).add(e)

            matched_projects = projects_found & icb_projects
            unmatched_projects = projects_found - icb_projects

            print(f"  Forecast Projects : {len(projects_found)}")
            print(f"  ICB Projects      : {len(icb_projects)}")
            print(f"  Matched Projects  : "
                  f"{len(matched_projects)}/{len(projects_found)}")
            if matched_projects:
                for p in sorted(matched_projects):
                    print(f"    [MATCH]    {p}")
            if unmatched_projects:
                print(f"  [WARNING] {len(unmatched_projects)} Forecast project(s) "
                      f"NOT found in ICB:")
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
                            f"  [NO MATCH] Consultant raw='{raw_id}', "
                            f"normalized='{norm_id}' in project "
                            f"'{proj_norm}' — missing from ICB"
                        )

            total_cons = total_matched_cons + total_unmatched_cons
            print(f"  Matched Consultants : "
                  f"{total_matched_cons}/{total_cons}")
            if total_unmatched_cons:
                print(f"  [WARNING] {total_unmatched_cons} consultant(s) "
                      f"not matched")

            # =============================================================
            # 6. FINAL RESULT
            # =============================================================
            cells_filled = len(self.forecast_data) * 3

            print("\n--- 6. FINAL RESULT ---")
            print(f"  Headers detected     : row {header_row_idx}")
            print(f"  Identity columns     : Project Num="
                  f"{_col_letter(col_proj)}, ICB ID="
                  f"{_col_letter(col_icb_id)}")
            print(f"  Monthly blocks       : {len(month_blocks)}")
            print(f"  Consultant rows      : {rows_processed}")
            print(f"  Rows skipped         : {rows_skipped}")
            print(f"  Forecast entries     : {len(self.forecast_data)}")
            print(f"  Cells filled         : {cells_filled} "
                  f"(NCR + Working Days + Cost)")
            print(
                f"\n[FORECAST IMPORT] {len(projects_found)} projects, "
                f"{total_matched_cons}/{total_cons} consultants matched, "
                f"{cells_filled} cells filled"
            )
            # =============================================================
            # 7. READING THE "EDM" SHEET
            # =============================================================
            print("\n--- 7. READING THE 'EDM' SHEET ---")
            edm_sheet = None
            for sname in wb.sheetnames:
                if sname.strip().lower() == "edm":
                    edm_sheet = wb[sname]
                    break

            self.project_em_map = {}
            if edm_sheet is not None:
                from openpyxl.utils import get_column_letter

                print(f"  Sheet 'EDM' detected: '{edm_sheet.title}'")

                # 7a. Scan the first 20 rows to detect the header row containing
                #     "Engagement Manager" (or typo) and at least one "Project Num"
                edm_header_row = None
                for r in range(1, min((edm_sheet.max_row or 1) + 1, 21)):
                    row_vals = [
                        edm_sheet.cell(r, c).value
                        for c in range(1, min((edm_sheet.max_column or 1) + 1, 50))
                    ]
                    has_em = False
                    has_proj = False
                    for val in row_vals:
                        if val and isinstance(val, str):
                            vc = val.strip().lower()
                            if vc in ("engagement manager", "engagemenet manager"):
                                has_em = True
                            elif vc in ("project num", "project number"):
                                has_proj = True
                    if has_em and has_proj:
                        edm_header_row = r
                        break

                if edm_header_row is not None:
                    # 7b. Detect columns in the header row
                    proj_num_cols = []
                    em_cols = []
                    for c in range(1, min((edm_sheet.max_column or 1) + 1, 50)):
                        val = edm_sheet.cell(edm_header_row, c).value
                        if val and isinstance(val, str):
                            vc = val.strip().lower()
                            if vc in ("project num", "project number"):
                                proj_num_cols.append(c)
                            elif vc in ("engagement manager", "engagemenet manager"):
                                em_cols.append(c)

                    # Project Num: must use column C (1-indexed 3).
                    # Among "Project Num" columns, take the second occurrence (the right one).
                    # If only 1 occurrence or second is not column C, fallback to column C with [WARNING].
                    col_proj = None
                    if len(proj_num_cols) >= 2:
                        second_col = proj_num_cols[1]
                        if second_col == 3:
                            col_proj = 3
                        else:
                            print(f"  [WARNING] [EDM] Second 'Project Num' column found at column {get_column_letter(second_col)} instead of C. Using column C by position.")
                            col_proj = 3
                    elif len(proj_num_cols) == 1:
                        if proj_num_cols[0] == 3:
                            col_proj = 3
                        else:
                            print(f"  [WARNING] [EDM] Only one 'Project Num' column found (column {get_column_letter(proj_num_cols[0])}). Using column C by position.")
                            col_proj = 3
                    else:
                        print("  [WARNING] [EDM] No 'Project Num' column found by name. Using column C by position.")
                        col_proj = 3

                    # Never retain column A (1) for joining
                    if col_proj == 1:
                        print("  [WARNING] [EDM] Column A detected for Project Num — overriding to column C.")
                        col_proj = 3

                    # Engagement Manager: column E (1-indexed 5).
                    # Detected by name with typo tolerance; fallback to column E with [WARNING].
                    col_em = None
                    if em_cols:
                        col_em = em_cols[0]
                    else:
                        print("  [WARNING] [EDM] 'Engagement Manager' column not found by name. Using column E by position.")
                        col_em = 5

                    col_proj_letter = get_column_letter(col_proj)
                    col_em_letter = get_column_letter(col_em)
                    print(f"  [EDM] Project Num = column {col_proj_letter}, Engagement Manager = column {col_em_letter}, column A ignored")

                    # 7c. Read rows
                    distinct_edm_projects = set()
                    for r in range(edm_header_row + 1, (edm_sheet.max_row or edm_header_row) + 1):
                        p_val = edm_sheet.cell(r, col_proj).value
                        em_val = edm_sheet.cell(r, col_em).value
                        if p_val is None or em_val is None:
                            continue
                        p_str = str(p_val).strip()
                        em_str = str(em_val).strip()
                        if not p_str or not em_str:
                            continue

                        norm_p = ExpenditureItem._normalize_key(p_str)
                        if not norm_p:
                            continue

                        distinct_edm_projects.add(norm_p)

                        if norm_p in self.project_em_map:
                            existing_em = self.project_em_map[norm_p]
                            if existing_em.lower() != em_str.lower():
                                print(f"  [WARNING] Duplicate Project Code in 'EDM': '{norm_p}' with different EM ('{existing_em}' vs '{em_str}'). Keeping first value.")
                        else:
                            self.project_em_map[norm_p] = em_str

                    # 7d. Diagnostic logs
                    loaded_projects = icb_projects | projects_found
                    matched_edm = distinct_edm_projects & loaded_projects
                    unmatched_edm = distinct_edm_projects - loaded_projects

                    print(f"  Distinct EDM projects                       : {len(distinct_edm_projects)}")
                    print(f"  EDM projects matched with loaded projects   : {len(matched_edm)}")
                    print(f"  EDM projects without match                  : {len(unmatched_edm)}")
                    print(f"  Projects with EM loaded                     : {len(self.project_em_map)}")
                else:
                    print("  [ERROR] [EDM] Unable to detect header row containing 'Engagement Manager' and 'Project Num' in the first 20 rows of 'EDM'.")

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
                    # Ignore entries whose work/cost values are 0 / empty
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

                # Do not inject if all activity values are zero
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

            # Recompute average NCR for projects (only on active consultants)
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

            # If the project has no active consultant and no expenditure or forecast, do not display it
            if not consultants_list and cc["totals"]["working_days"] == 0 and cc["totals"]["bill_amount"] == 0 and cc["totals"]["forecast_bill_amount"] == 0:
                continue

            norm_proj = ExpenditureItem._normalize_key(cc_proj)
            em_name = getattr(self, "project_em_map", {}).get(norm_proj, "")

            result.append({
                "cc_project_number": cc_proj,
                "project_name": cc["project_name"],
                "em_name": em_name,
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

    def _generate_gap_comment(
        self,
        cc,
        norm_proj,
        month,
        forecast_ba_total,
        has_any_forecast,
        gap_tolerance=2.0,
    ):
        """Generate automatic explanatory comment for gaps on a project."""
        comment = ""
        icb_ba = cc['totals']['bill_amount']
        gap_val = round(icb_ba - forecast_ba_total, 2) if has_any_forecast else 0

        if not (has_any_forecast and abs(gap_val) > gap_tolerance):
            return ""

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
                if name_str.startswith("(") and name_str.endswith(")"):
                    name_str = name_str[1:-1].strip()
                return f"{ne} ({name_str})"
            return f"{ne}" if ne else "(Unknown)"

        consultant_blocks = []

        for ne in sorted(icb_emps.keys()):
            cons = icb_emps[ne]
            fk = (norm_proj, ne, month)
            fd = self.forecast_data.get(fk)
            if not fd:
                continue

            gaps = []

            # Working days comparison: strict equality
            wd_icb = round(cons['working_days'], 2)
            wd_fc = round(fd['working_days'], 2)
            if wd_icb != wd_fc:
                gaps.append(f"Gap due to working days: ICB={wd_icb} D/Forecast={wd_fc} D")

            # NCR comparison: uses dynamic tolerance
            ncr_icb = round(cons['ncr'], 2)
            ncr_fc = round(fd['ncr'], 2)
            if abs(ncr_icb - ncr_fc) > gap_tolerance:
                gaps.append(f"Gap due to rate (NCR): ICB={ncr_icb}/Forecast={ncr_fc}")

            # Cost comparison: uses dynamic tolerance
            ba_icb = round(cons['total_bill_amount'], 2)
            ba_fc = round(fd['bill_amount'], 2)
            if abs(ba_icb - ba_fc) > gap_tolerance:
                gaps.append(f"Cost issue: ICB={ba_icb}/Forecast={ba_fc}")

            if gaps:
                label = _id_name(ne)
                block = f"{label}:\n" + "\n".join(gaps)
                consultant_blocks.append(block)

        if consultant_blocks:
            comment = "\n\n".join(consultant_blocks)
        else:
            comment = "Cost issue (requires manual verification)"

        return comment

    def export_to_excel(
        self,
        month_filter,
        start_date,
        end_date,
        gap_tolerance=2.0,
        include_comments=True,
        include_working_days=False,
        include_raw_cost=False,
        include_controlling_comments=True,
    ):
        import io
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.worksheet.table import Table, TableStyleInfo
        from openpyxl.utils import get_column_letter

        # 1. Get items in the range to determine the months involved
        items_in_range = self._get_items(month_filter=month_filter, start_date=start_date, end_date=end_date)
        months_in_range = sorted(list(set(i.month_key for i in items_in_range)))

        wb = Workbook()
        # Remove the default sheet if we will create specific sheets
        default_ws = wb.active

        # Styles
        right_align = Alignment(horizontal="right")
        wrap_align = Alignment(wrap_text=True)

        if not months_in_range:
            default_ws.title = "No data"
            default_ws.append(["No data in this range"])
        else:
            for idx, month in enumerate(months_in_range):
                if idx == 0:
                    ws = default_ws
                    ws.title = month
                else:
                    ws = wb.create_sheet(title=month)

                data = self.hierarchy_data(month_filter=month, start_date=start_date, end_date=end_date)

                # Column specifications: (key, header_title, width, is_numeric, is_comment)
                col_specs = [
                    ("em_lead", "EM_Lead", 25, False, False),
                    ("project_name", "Project Name", 40, False, False),
                    ("cc_project_number", "CC Project Number", 25, False, False),
                ]
                if include_working_days:
                    col_specs.append(("working_days", "Working Days", 15, True, False))

                col_specs.append(("icb_tp", "ICB Transfer Price", 20, True, False))
                col_specs.append(("forecast_tp", "Forecast Transfer Price", 22, True, False))
                col_specs.append(("gap", "Gap", 15, True, False))

                if include_raw_cost:
                    col_specs.append(("raw_cost", "Raw Cost", 15, True, False))

                col_specs.append(("issue_type", "Issue Type", 25, False, False))

                if include_comments:
                    col_specs.append(("generated_comment", "Generated comment", 60, False, True))

                if include_controlling_comments:
                    col_specs.append(("controlling_comment", "Controlling comment", 40, False, True))

                headers = [spec[1] for spec in col_specs]
                num_cols = len(headers)
                ws.append(headers)

                icb_col_idx = [i + 1 for i, s in enumerate(col_specs) if s[0] == "icb_tp"][0]
                fc_col_idx = [i + 1 for i, s in enumerate(col_specs) if s[0] == "forecast_tp"][0]
                icb_col_letter = get_column_letter(icb_col_idx)
                fc_col_letter = get_column_letter(fc_col_idx)

                current_row = 2
                for cc in data:
                    norm_proj = ExpenditureItem._normalize_key(cc['cc_project_number'])

                    # --- Forecast Transfer Price column ---
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
                    if include_comments:
                        comment = self._generate_gap_comment(
                            cc, norm_proj, month, forecast_ba_total, has_any_forecast, gap_tolerance
                        )

                    gap_formula = f"={icb_col_letter}{current_row}-{fc_col_letter}{current_row}"
                    em_lead_val = cc.get("em_name") or getattr(self, "project_em_map", {}).get(norm_proj, "")

                    value_map = {
                        "em_lead": em_lead_val,
                        "project_name": cc['project_name'] or "",
                        "cc_project_number": cc['cc_project_number'],
                        "working_days": cc['totals']['working_days'],
                        "icb_tp": cc['totals']['bill_amount'],
                        "forecast_tp": forecast_cell_value,
                        "gap": gap_formula,
                        "raw_cost": cc['totals']['raw_cost'],
                        "issue_type": "",
                        "generated_comment": comment if include_comments else "",
                        "controlling_comment": "",
                    }
                    row_data = [value_map[spec[0]] for spec in col_specs]
                    ws.append(row_data)

                    # Alignment
                    for col in range(1, num_cols + 1):
                        cell = ws.cell(row=current_row, column=col)
                        spec = col_specs[col - 1]
                        if spec[3]:  # is_numeric
                            cell.alignment = right_align
                        elif spec[4]:  # is_comment
                            cell.alignment = Alignment(
                                wrap_text=True,
                                vertical='top',
                                horizontal='left',
                            )
                            if spec[0] == "generated_comment" and comment:
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
                for col_idx, spec in enumerate(col_specs, start=1):
                    ws.column_dimensions[get_column_letter(col_idx)].width = spec[2]

            # =================================================================
            # KPI DASHBOARD SHEET
            # =================================================================
            self._build_kpi_sheet(wb, months_in_range, start_date, end_date, gap_tolerance)

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    def export_to_csv(
        self,
        month_filter,
        start_date,
        end_date,
        gap_tolerance=2.0,
        include_comments=True,
        include_working_days=False,
        include_raw_cost=False,
        include_controlling_comments=True,
    ):
        import io
        import csv

        items_in_range = self._get_items(month_filter=month_filter, start_date=start_date, end_date=end_date)
        months_in_range = sorted(list(set(i.month_key for i in items_in_range)))

        output = io.StringIO()
        writer = csv.writer(output, delimiter=',', quoting=csv.QUOTE_MINIMAL)

        col_specs = [
            ("month", "Month"),
            ("em_lead", "EM_Lead"),
            ("project_name", "Project Name"),
            ("cc_project_number", "CC Project Number"),
        ]
        if include_working_days:
            col_specs.append(("working_days", "Working Days"))

        col_specs.append(("icb_tp", "ICB Transfer Price"))
        col_specs.append(("forecast_tp", "Forecast Transfer Price"))
        col_specs.append(("gap", "Gap"))

        if include_raw_cost:
            col_specs.append(("raw_cost", "Raw Cost"))

        col_specs.append(("issue_type", "Issue Type"))

        if include_comments:
            col_specs.append(("generated_comment", "Generated comment"))

        if include_controlling_comments:
            col_specs.append(("controlling_comment", "Controlling comment"))

        headers = [spec[1] for spec in col_specs]
        writer.writerow(headers)

        for month in months_in_range:
            data = self.hierarchy_data(month_filter=month, start_date=start_date, end_date=end_date)
            for cc in data:
                norm_proj = ExpenditureItem._normalize_key(cc['cc_project_number'])

                has_any_forecast = False
                forecast_ba_total = 0.0
                for fk, fd in self.forecast_data.items():
                    if fk[0] == norm_proj and fk[2] == month:
                        if abs(fd.get("working_days", 0)) > 1e-6 or abs(fd.get("bill_amount", 0)) > 1e-6:
                            has_any_forecast = True
                            forecast_ba_total += fd["bill_amount"]

                forecast_cell_value = round(forecast_ba_total, 2) if has_any_forecast else ""

                comment = ""
                if include_comments:
                    comment = self._generate_gap_comment(
                        cc, norm_proj, month, forecast_ba_total, has_any_forecast, gap_tolerance
                    )

                em_lead_val = cc.get("em_name") or getattr(self, "project_em_map", {}).get(norm_proj, "")
                icb_ba = cc['totals']['bill_amount']
                gap_val = round(icb_ba - forecast_ba_total, 2) if has_any_forecast else ""

                value_map = {
                    "month": month,
                    "em_lead": em_lead_val,
                    "project_name": cc['project_name'] or "",
                    "cc_project_number": cc['cc_project_number'],
                    "working_days": cc['totals']['working_days'],
                    "icb_tp": icb_ba,
                    "forecast_tp": forecast_cell_value,
                    "gap": gap_val,
                    "raw_cost": cc['totals']['raw_cost'],
                    "issue_type": "",
                    "generated_comment": comment if include_comments else "",
                    "controlling_comment": "",
                }
                row_data = [value_map[spec[0]] for spec in col_specs]
                writer.writerow(row_data)

        csv_bytes = output.getvalue().encode("utf-8-sig")
        buffer = io.BytesIO(csv_bytes)
        buffer.seek(0)
        return buffer

    def _build_kpi_sheet(self, wb, months_in_range, start_date, end_date, gap_tolerance):
        """Build the KPI Dashboard sheet in the export workbook.
        
        Sections:
        1. Monthly summary (Cost, Forecast, Gap, WD, Raw Cost, Nb projects, Nb consultants)
        2. Cumulative Bill Amount by month (month-by-month progression)
        3. Summary by project (all periods combined)
        4. Top consultants by Cost
        5. Consultant detail by month
        """
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
        from openpyxl.utils import get_column_letter

        ws = wb.create_sheet(title="KPI Dashboard")

        # ---- Reusable styles ----
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

        # ---- Collect all data by month ----
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

                # Forecast for this project/month
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

                # Aggregate by project
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

                    # Monthly consultant detail
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
        # SECTION 0: MAIN TITLE
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
            period_label = f"Period: {start_date} to {end_date}"
        elif months_in_range:
            period_label = f"Months: {months_in_range[0]} to {months_in_range[-1]}"
        sub_cell = ws.cell(row=row, column=1, value=period_label)
        sub_cell.font = Font(name="Calibri", size=10, italic=True, color="8899AA")
        sub_cell.fill = dark_fill
        for c in range(1, 10):
            ws.cell(row=row, column=c).fill = dark_fill
        row += 2

        # ================================================================
        # SECTION 1: MONTHLY SUMMARY
        # ================================================================
        _write_section_title(ws, row, "1. MONTHLY SUMMARY", 8)
        row += 1
        headers_s1 = ["Month", "ICB Transfer Price", "Forecast Transfer Price", "Gap (ICB-Forecast)",
                       "Working Days", "Raw Cost", "Nb Projects", "Nb Consultants"]
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

        # TOTAL row
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
        # SECTION 2: CUMULATIVE TRANSFER PRICE BY MONTH
        # ================================================================
        _write_section_title(ws, row, "2. CUMULATIVE TRANSFER PRICE BY MONTH (Progression)", 8)
        row += 1
        headers_s2 = ["Month", "Monthly ICB TP", "Cumulative ICB TP",
                       "Monthly Forecast TP", "Cumulative Forecast TP",
                       "Cumulative Gap", "TP Progression %"]
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
        # SECTION 3: SUMMARY BY PROJECT
        # ================================================================
        _write_section_title(ws, row, "3. SUMMARY BY PROJECT (all periods)", 9)
        row += 1
        headers_s3 = ["CC Project Number", "Project Name", "ICB TP Total",
                       "Forecast TP Total", "Gap", "Gap %",
                       "Working Days", "Avg NCR", "Nb Consultants"]
        for c, h in enumerate(headers_s3, 1):
            ws.cell(row=row, column=c, value=h)
        _style_header_row(ws, row, 1, len(headers_s3))
        row += 1

        # Sort by cost descending
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

            # Alternating row colors
            if (row % 2) == 0:
                for c in range(1, 10):
                    if ws.cell(row=row, column=c).fill == PatternFill():
                        ws.cell(row=row, column=c).fill = light_row_fill

            row += 1

        row += 1

        # ================================================================
        # SECTION 4: TOP CONSULTANTS BY COST
        # ================================================================
        _write_section_title(ws, row, "4. TOP CONSULTANTS BY TRANSFER PRICE (all periods)", 8)
        row += 1
        headers_s4 = ["Employee ID", "Name", "Main Project",
                       "ICB TP Total", "Working Days Total",
                       "NCR", "Active Months", "Avg TP / Month"]
        for c, h in enumerate(headers_s4, 1):
            ws.cell(row=row, column=c, value=h)
        _style_header_row(ws, row, 1, len(headers_s4))
        row += 1

        # Aggregate by consultant
        cons_totals = []
        for (emp_id, name), months_data in consultant_monthly.items():
            total_cost = sum(md["cost"] for md in months_data.values())
            total_wd = sum(md["wd"] for md in months_data.values())
            ncr_vals = [md["ncr"] for md in months_data.values() if md["ncr"] > 0]
            avg_ncr = round(sum(ncr_vals) / len(ncr_vals), 2) if ncr_vals else 0
            # Most frequent project
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

        # Sort by cost descending, limit to 50
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
        # SECTION 5: CONSULTANT DETAIL BY MONTH
        # ================================================================
        _write_section_title(ws, row,
                             "5. CONSULTANT DETAIL BY MONTH (TP per consultant per month)",
                             3 + len(months_in_range))
        row += 1

        # Headers: Employee ID | Name | Project | Month1 | Month2 | ... | TOTAL
        headers_s5 = ["Employee ID", "Name", "Project"] + months_in_range + ["TOTAL"]
        for c, h in enumerate(headers_s5, 1):
            ws.cell(row=row, column=c, value=h)
        _style_header_row(ws, row, 1, len(headers_s5))
        row += 1

        # Sort by name
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

            # TOTAL column
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

        # ---- Column widths ----
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['B'].width = 35
        ws.column_dimensions['C'].width = 22
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 18
        ws.column_dimensions['F'].width = 18
        ws.column_dimensions['G'].width = 18
        ws.column_dimensions['H'].width = 18
        ws.column_dimensions['I'].width = 18

        # Freeze the first row
        ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Flask Routes
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload_file():
    global pipeline
    if "file" not in request.files:
        return jsonify({"error": "No file provided."}), 400
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "File must be in .xlsx format"}), 400

    # Temporary save
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
        return jsonify({"error": f"Error during loading: {str(exc)}"}), 500

    return jsonify({
        "total_lines": len(pipeline.items),
        "filtered_lines": len(pipeline.filtered_items),
        "available_months": pipeline.available_months,
    })

@app.route("/api/upload_forecast", methods=["POST"])
def upload_forecast():
    if "file" not in request.files:
        return jsonify({"error": "No file provided."}), 400
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "File must be in .xlsx format"}), 400

    # Temporary save
    upload_dir = os.path.join(os.path.dirname(__file__), "uploads")
    os.makedirs(upload_dir, exist_ok=True)
    path = os.path.join(upload_dir, "current_forecast.xlsx")
    f.save(path)

    try:
        pipeline.load_forecast(path)
        return jsonify({"message": "Forecast file loaded successfully."}), 200
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/reset", methods=["POST"])
def reset_data():
    """Reset the application: clear the pipeline and delete temporary files."""
    global pipeline
    pipeline = None

    # Delete uploaded files
    upload_dir = os.path.join(os.path.dirname(__file__), "uploads")
    for fname in ("current_icb.xlsx", "current_forecast.xlsx"):
        fpath = os.path.join(upload_dir, fname)
        if os.path.exists(fpath):
            try:
                os.remove(fpath)
            except OSError:
                pass

    return jsonify({"message": "Data unloaded successfully."})


@app.route("/api/hierarchy")
def get_hierarchy():
    global pipeline
    if pipeline is None:
        return jsonify({"error": "No file loaded."}), 400
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
        return jsonify({"error": "No file loaded."}), 400
    start_date = request.args.get("start_date", None)
    end_date = request.args.get("end_date", None)
    items = pipeline._get_items(start_date=start_date, end_date=end_date)
    months = sorted(list(set(i.month_key for i in items)))
    return jsonify(months)


@app.route("/api/transactions")
def get_transactions():
    global pipeline
    if pipeline is None:
        return jsonify({"error": "No file loaded."}), 400
    emp = request.args.get("employee")
    cc_proj = request.args.get("cc_project")
    month = request.args.get("month", None)
    if month == "all" or month == "":
        month = None
    start_date = request.args.get("start_date", None)
    end_date = request.args.get("end_date", None)
    
    if not emp or not cc_proj:
        return jsonify({"error": "Parameters employee and cc_project are required."}), 400
        
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
        return jsonify({"error": "No file loaded."}), 400
    
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
    
    export_format = request.args.get("format", "xlsx").lower().strip()
    if export_format not in ("csv", "xlsx"):
        export_format = "xlsx"

    # Whether to include optional columns
    include_working_days = request.args.get("include_working_days", "0") in ("1", "true", "yes")
    include_raw_cost = request.args.get("include_raw_cost", "0") in ("1", "true", "yes")
    include_comments = request.args.get("include_comments", "1") in ("1", "true", "yes")
    include_controlling_comments = request.args.get("include_controlling_comments", "1") in ("1", "true", "yes")
    
    ext = "csv" if export_format == "csv" else "xlsx"

    try:
        if export_format == "csv":
            buffer = pipeline.export_to_csv(
                month_filter=month,
                start_date=start_date,
                end_date=end_date,
                gap_tolerance=gap_tolerance,
                include_comments=include_comments,
                include_working_days=include_working_days,
                include_raw_cost=include_raw_cost,
                include_controlling_comments=include_controlling_comments,
            )
        else:
            buffer = pipeline.export_to_excel(
                month_filter=month,
                start_date=start_date,
                end_date=end_date,
                gap_tolerance=gap_tolerance,
                include_comments=include_comments,
                include_working_days=include_working_days,
                include_raw_cost=include_raw_cost,
                include_controlling_comments=include_controlling_comments,
            )
        
        import os
        if start_date and end_date:
            filename = f"icb_{start_date}_{end_date}.{ext}"
        elif month:
            filename = f"icb_{month}.{ext}"
        else:
            filename = f"icb_export.{ext}"
        filepath = None
        
        file_filters = ('CSV Files (*.csv)', 'All Files (*.*)') if export_format == "csv" else ('Excel Files (*.xlsx)', 'All Files (*.*)')

        try:
            import webview
            if webview.windows:
                window = webview.windows[0]
                result = window.create_file_dialog(
                    webview.FileDialog.SAVE,
                    save_filename=filename,
                    file_types=file_filters
                )
                if result and isinstance(result, (tuple, list)) and len(result) > 0:
                    filepath = result[0]
                elif result and isinstance(result, str):
                    filepath = result
                else:
                    return jsonify({"error": "Export cancelled."}), 400
        except ImportError:
            pass
            
        if not filepath:
            # Build the desktop path as fallback
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            filepath = os.path.join(desktop, filename)
        elif not filepath.lower().endswith(f".{ext}"):
            filepath = f"{filepath}.{ext}"
        
        with open(filepath, "wb") as f:
            f.write(buffer.getvalue())
            
        return jsonify({
            "success": True,
            "message": f"{export_format.upper()} export successful",
            "path": filepath
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def start_server(port=5050):
    """Start the Flask server (used by the desktop wrapper)."""
    import logging
    log = logging.getLogger("werkzeug")
    log.setLevel(logging.ERROR)
    app.run(debug=True, port=port, use_reloader=False)

    