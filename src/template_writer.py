from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.comments import Comment
from copy import copy
import re
import pandas as pd
from src.models import CostCenter, WBSCode, PO, MonthlyMetrics, ExceptionLog, ExceptionType


def month_sort_key(month_str):
    """Convert month string to sortable key for proper chronological ordering"""
    month_order = {
        'Dec (PY)': 0,
        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12,
        'Unknown': 13
    }
    return month_order.get(month_str, 13)



class TemplateWriter:

    def __init__(self, 
                 file_path,
                 output_path,
                 overwrite,
                 header_row, 
                 po_column,
                 dec_acc_reversal_col, 
                 forecast_source_cols,
                 transactional_source_cols
        ):
        
        self.wb = load_workbook(file_path)
        self.sheet: Worksheet = self.wb.active  # type: ignore[assignment]
        if self.sheet is None:
            raise ValueError(f"Could not load active sheet from {file_path}")

        self.output_path = output_path

        # Overwrites previous months data if true
        self.overwrite = overwrite

        # Configs for template
        self.header_row = header_row    # Header row
        self.po_column = po_column    # Col where POs are entered


        # Cost centers, WBS codes, PO numbers, and their associated rows
        self.cost_centers = {}
        self.wbs_codes = {}

        # Column map (dynamically created starting with Dec Accrual Reversal)
        self.dec_acc_reversal_col = dec_acc_reversal_col # Col where first data entry exists
        self.column_map = self.get_column_map(starting_col=self.dec_acc_reversal_col)

        # Source sheet params
        self.forecast_source_cols = forecast_source_cols
        self.transactional_source_cols = transactional_source_cols

        # PO columns
        self.forecast_po_col = self.forecast_source_cols[0]
        self.transactional_po_col = self.transactional_source_cols[0]

        # Total 2026 column — found dynamically by scanning the header row
        self.total_col = self._get_total_col()

    @staticmethod
    def _norm_po(v) -> str:
        """Normalise a PO value to a clean string (strip .0 from float-integers)."""
        s = str(v).strip()
        if s.replace('.', '', 1).replace('-', '', 1).isdigit():
            try:
                return str(int(float(s)))
            except (ValueError, OverflowError):
                pass
        return s

    def _get_total_col(self) -> str | None:
        """Scan all rows up to and including the first data row for a cell containing
        'Total' followed by a year (e.g. 'Total 2026'). Returns the column letter,
        or None if not found."""
        search_rows = range(1, self.header_row + 10)  # generous window above data
        for row in search_rows:
            for col in range(1, (self.sheet.max_column or 200) + 1):
                val = self.sheet.cell(row=row, column=col).value
                if val and re.search(r'total\s+\d{4}', str(val), re.IGNORECASE):
                    return get_column_letter(col)
        return None

    def _write_total_formula(self, row: int):
        """Write the Total 2026 SUM formula into the total column for the given row.
        The formula picks the best available value per month:
        Actual → Accrual → Forecast (matching the pattern in the template).
        Only writes if the total column was found and the cell is blank (or overwrite=True)."""
        if not self.total_col:
            return
        cell = self.sheet[f"{self.total_col}{row}"]
        if not self.overwrite and cell.value not in (None, 0, ''):
            return
        # Build SUM of per-month best-value: IF(Actual<>"", Actual, IF(Accrual<>"", Accrual, Forecast))
        parts = []
        for month_cols in self.column_map.values():
            actual   = month_cols.get('Actual')
            accrual  = month_cols.get('Accrual')
            forecast = month_cols.get('Forecast')
            if actual and accrual and forecast:
                parts.append(
                    f'IF({actual}{row}<>"",{actual}{row},'
                    f'IF({accrual}{row}<>"",{accrual}{row},{forecast}{row}))'
                )
        if parts:
            cell.value = '=' + '+'.join(parts)

    # Map full/variant month words found in header cells → canonical 3-letter key
    _HEADER_MONTH_ALIASES = {
        "jan": "Jan", "feb": "Feb",
        "mar": "Mar", "march": "Mar",
        "apr": "Apr", "april": "Apr",
        "may": "May",
        "jun": "Jun", "june": "Jun",
        "jul": "Jul", "july": "Jul",
        "aug": "Aug",
        "sep": "Sep", "sept": "Sep",
        "oct": "Oct", "nov": "Nov", "dec": "Dec",
    }

    def get_column_map(self, starting_col):
        """Build month→metric→column-letter map by scanning the template header row.

        The header row contains cells like 'Forecast Jan', 'Accrual Reversal Dec',
        'Actual   Feb', etc.  We parse each cell to extract the metric keyword and
        the month keyword, normalise both, and record the column letter.

        The first 'Accrual Reversal Dec' column (col N in the current template) is
        the prior-year December accrual reversal — mapped to the key 'Dec (PY)' so
        it stays distinct from the current-year 'Dec' block.

        Falls back to the old fixed-offset arithmetic if no header cells are found
        (e.g. a blank template that hasn't been opened in Excel yet).
        """
        metrics_lower = {
            "accrual reversal": "Accrual Reversal",
            "forecast":         "Forecast",
            "accrual":          "Accrual",
            "actual":           "Actual",
        }

        col_map: dict[str, dict[str, str]] = {}
        dec_py_seen = False  # first 'Dec' Accrual Reversal is Dec (PY)
        max_col = self.sheet.max_column or 200

        for col_idx in range(1, max_col + 1):
            raw = self.sheet.cell(row=self.header_row, column=col_idx).value
            if not raw:
                continue
            text = str(raw).strip()

            # Identify which metric this cell represents
            matched_metric = None
            for kw, canonical in metrics_lower.items():
                if text.lower().startswith(kw):
                    matched_metric = canonical
                    break
            if matched_metric is None:
                continue

            # Extract the month word (last token in the cell text)
            words = text.split()
            month_word = words[-1].lower().rstrip('.,').strip() if words else ""
            month_key = self._HEADER_MONTH_ALIASES.get(month_word)
            if month_key is None:
                continue

            # The very first 'Accrual Reversal Dec' is the prior-year Dec column
            if matched_metric == "Accrual Reversal" and month_key == "Dec" and not dec_py_seen:
                month_key = "Dec (PY)"
                dec_py_seen = True

            if month_key not in col_map:
                col_map[month_key] = {}
            col_map[month_key][matched_metric] = get_column_letter(col_idx)

        if col_map:
            return col_map

        # ----------------------------------------------------------------
        # Fallback: old fixed-offset arithmetic (used when the template has
        # no populated header row, e.g. a truly blank workbook).
        # ----------------------------------------------------------------
        print(
            "WARNING: No metric headers found in header row — "
            "falling back to fixed column offsets starting at column "
            f"'{starting_col}'."
        )
        months_fb = [
            "Dec (PY)",
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ]
        metrics_fb = ["Accrual Reversal", "Forecast", "Accrual", "Actual"]
        col_index = column_index_from_string(starting_col)
        col_map_fb: dict[str, dict[str, str]] = {}
        for month in months_fb:
            month_map: dict[str, str] = {}
            for metric in metrics_fb:
                month_map[metric] = get_column_letter(col_index)
                col_index += 1
            col_map_fb[month] = month_map
            col_index += 1  # skip variance column
        return col_map_fb

    def insert_missing_po_rows(self, hierarchy: dict, pos: dict[str, int]) -> dict[str, int]:
        """
        For each cost center in the hierarchy:
        - If the cost center already has PO rows in the template, insert only the
          rows that are missing (including RECLASS) directly after its last existing row.
        - If the cost center has no rows at all, insert all PO rows (including RECLASS)
          just above the 'Previous Period Invoices' stop marker.
        Existing written rows are never modified.
        """
        stop_row = None
        for row_idx in range(1, (self.sheet.max_row or 1000) + 1):
            cell_val = self.sheet[f"A{row_idx}"].value
            if cell_val is not None and str(cell_val).strip() == "Previous Period Invoices":
                stop_row = row_idx
                break

        if stop_row is None:
            print("WARNING: 'Previous Period Invoices' marker not found – PO rows not inserted.")
            return pos

        po_col_idx = column_index_from_string(self.po_column)
        cc_col_idx = 10
        wbs_col_idx = 6  # Column F
        ref_row = self.header_row + 1
        max_col = self.sheet.max_column or 1

        ref_styles = {}
        for col_idx in range(1, max_col + 1):
            src = self.sheet.cell(row=ref_row, column=col_idx)
            ref_styles[col_idx] = {
                'font': copy(src.font),
                'border': copy(src.border),
                'alignment': copy(src.alignment),
                'number_format': src.number_format,
                'fill': copy(src.fill),
            }
        ref_height = self.sheet.row_dimensions[ref_row].height

        def _scan_existing_rows_by_cc() -> dict:
            """Re-scan the sheet to get the last existing row per cost center."""
            result = {}
            for row_idx in range(self.header_row + 1, (self.sheet.max_row or 1000) + 1):
                cell_val = self.sheet[f"A{row_idx}"].value
                if cell_val is not None and str(cell_val).strip() == "Previous Period Invoices":
                    break
                cc_val = self.sheet.cell(row=row_idx, column=cc_col_idx).value
                po_val = self.sheet.cell(row=row_idx, column=po_col_idx).value
                cc_text = str(cc_val).strip().split('/')[0].strip() if cc_val is not None and str(cc_val).strip() != "" else None
                po_text = str(po_val).strip() if po_val is not None and str(po_val).strip() != "" else None
                if cc_text and po_text:
                    result[cc_text] = row_idx  # keep updating → ends at last row for this CC
            return result

        inserted = []
        for cc_id, cost_center in hierarchy.items():
            # Collect (po_number, wbs_code) pairs for POs not yet in the template
            po_numbers = []
            po_wbs_map = {}      # po_number -> wbs_code
            po_le_map = {}       # po_number -> legal_entity
            po_country_map = {}  # po_number -> country
            po_vendor_map = {}   # po_number -> vendor_name
            po_gl_map = {}       # po_number -> gl_account
            po_gross_map = {}    # po_number -> gross_po_value
            po_req_map = {}      # po_number -> req_title
            for wbs_code, wbs in cost_center.wbs_codes.items():
                if wbs_code.upper() == "ER":
                    continue
                for po_number, po_obj in wbs.pos.items():
                    norm_po = self._norm_po(po_number)
                    if norm_po not in pos and norm_po not in po_wbs_map:
                        po_number = norm_po
                        po_numbers.append(po_number)
                        po_wbs_map[po_number] = wbs_code
                        po_le_map[po_number] = po_obj.legal_entity
                        po_country_map[po_number] = po_obj.country
                        po_vendor_map[po_number] = po_obj.vendor_name
                        po_gl_map[po_number] = po_obj.gl_account
                        po_gross_map[po_number] = po_obj.gross_po_value
                        po_req_map[po_number] = po_obj.req_title

            if not po_numbers:
                continue

            # Re-scan each time so row numbers stay correct after previous insertions
            existing_last_row_by_cc = _scan_existing_rows_by_cc()

            if cc_id in existing_last_row_by_cc:
                # CC already has rows — insert new rows immediately after the last one
                insert_at = existing_last_row_by_cc[cc_id] + 1
            else:
                # CC has no rows yet — insert just above the stop marker
                # Re-find stop_row since it shifts with every insertion
                insert_at = None
                for row_idx in range(1, (self.sheet.max_row or 1000) + 1):
                    if self.sheet[f"A{row_idx}"].value is not None and \
                            str(self.sheet[f"A{row_idx}"].value).strip() == "Previous Period Invoices":
                        insert_at = row_idx
                        break
                if insert_at is None:
                    print("WARNING: stop marker disappeared during insertion.")
                    break

            for po_number in po_numbers:
                self.sheet.insert_rows(insert_at)
                for col_idx in range(1, max_col + 1):
                    new_cell = self.sheet.cell(row=insert_at, column=col_idx)
                    s = ref_styles[col_idx]
                    new_cell.font = copy(s['font'])
                    new_cell.border = copy(s['border'])
                    new_cell.alignment = copy(s['alignment'])
                    new_cell.number_format = s['number_format']
                    new_cell.fill = copy(s['fill'])
                if ref_height:
                    self.sheet.row_dimensions[insert_at].height = ref_height

                vendor_val = po_vendor_map.get(po_number)
                if vendor_val is not None:
                    self.sheet.cell(row=insert_at, column=5, value=vendor_val)
                self.sheet.cell(row=insert_at, column=cc_col_idx, value=cc_id)
                self.sheet.cell(row=insert_at, column=po_col_idx, value=po_number)
                self.sheet.cell(row=insert_at, column=wbs_col_idx, value=po_wbs_map.get(po_number))
                gl_val = po_gl_map.get(po_number)
                if gl_val is not None:
                    self.sheet.cell(row=insert_at, column=7, value=gl_val)
                le_val = po_le_map.get(po_number)
                if le_val is not None:
                    self.sheet.cell(row=insert_at, column=8, value=le_val)
                country_val = po_country_map.get(po_number)
                if country_val is not None:
                    self.sheet.cell(row=insert_at, column=9, value=country_val)
                gross_val = po_gross_map.get(po_number)
                if gross_val is not None:
                    self.sheet.cell(row=insert_at, column=11, value=round(gross_val, 2))
                req_val = po_req_map.get(po_number)
                if req_val is not None:
                    self.sheet.cell(row=insert_at, column=12, value=req_val)
                self._write_total_formula(insert_at)
                pos[po_number] = insert_at
                inserted.append(po_number)

                # Every insertion shifts all rows below by 1 — update pos accordingly
                for key in pos:
                    if key != po_number and pos[key] >= insert_at:
                        pos[key] += 1

                insert_at += 1

        if inserted:
            print(f"Inserted {len(inserted)} PO row(s): {inserted}")
        else:
            print("No additional PO rows needed.")
        return pos

    def insert_er_rows(self, hierarchy: dict, pos: dict[str, int]) -> dict[str, int]:
        """
        Collects all ER numbers from the hierarchy (those assigned WBS="ER" during
        exception processing) and inserts a new row for each one in the Template sheet
        directly above the 'Previous Period Invoices' stop-marker row.

        The new rows have only the ER number written into the PO column (column B by
        default, matching self.po_column).  All other cells are left blank so the
        normal write_hierarchy call can fill in the monthly data.

        Args:
            hierarchy: The built hierarchy dict returned by build_hierarchy.
            pos: The existing PO → row mapping from TemplateReader.get_existing_pos().

        Returns:
            Updated pos dict that includes the newly inserted ER rows.
        """
        # Collect unique ER numbers that are NOT already in the template
        er_pattern = re.compile(r'^ER\d+$', re.IGNORECASE)
        er_numbers = []
        er_cc_map = {}       # er_number -> cost_center_id
        er_le_map = {}       # er_number -> legal_entity
        er_country_map = {}  # er_number -> country
        er_vendor_map = {}   # er_number -> vendor_name
        er_gl_map = {}       # er_number -> gl_account
        er_gross_map = {}    # er_number -> gross_po_value
        er_req_map = {}      # er_number -> req_title
        er_wbs_map = {}      # er_number -> real WBS (po.real_wbs if set, else "ER")
        seen = set()
        for cc_id, cost_center in hierarchy.items():
            for wbs_code, wbs in cost_center.wbs_codes.items():
                if wbs_code.upper() == "ER":
                    for po_number, po_obj in wbs.pos.items():
                        if er_pattern.match(po_number) and po_number not in pos and po_number not in seen:
                            er_numbers.append(po_number)
                            er_cc_map[po_number] = cc_id
                            er_le_map[po_number] = po_obj.legal_entity
                            er_country_map[po_number] = po_obj.country
                            er_vendor_map[po_number] = po_obj.vendor_name
                            er_gl_map[po_number] = po_obj.gl_account
                            er_gross_map[po_number] = po_obj.gross_po_value
                            er_req_map[po_number] = po_obj.req_title
                            er_wbs_map[po_number] = po_obj.real_wbs or "ER"
                            seen.add(po_number)

        if not er_numbers:
            print("No ER rows to insert into template.")
            return pos

        # Find the stop-marker row (Previous Period Invoices) in the sheet
        stop_row = None
        for row_idx in range(1, (self.sheet.max_row or 1000) + 1):
            cell_val = self.sheet[f"A{row_idx}"].value
            if cell_val is not None and str(cell_val).strip() == "Previous Period Invoices":
                stop_row = row_idx
                break

        if stop_row is None:
            print("WARNING: 'Previous Period Invoices' marker not found – ER rows not inserted.")
            return pos

        # Update header label in the PO column to reflect ER numbers
        po_col_idx = column_index_from_string(self.po_column)
        header_cell = self.sheet.cell(row=self.header_row, column=po_col_idx)
        header_cell.value = "PO Number/ER Number"

        # Green fill for ER cells
        green_fill = PatternFill(fill_type="solid", fgColor="00B050")

        # Reference row: use the first data row (header_row + 1) as the style source.
        # This ensures we always copy from a well-formed row regardless of how many
        # ERs have already been inserted above stop_row.
        ref_row = self.header_row + 1
        max_col = self.sheet.max_column or 1

        # Build per-column style snapshots from the reference row once
        ref_styles = {}
        for col_idx in range(1, max_col + 1):
            src = self.sheet.cell(row=ref_row, column=col_idx)
            ref_styles[col_idx] = {
                'font': copy(src.font),
                'border': copy(src.border),
                'alignment': copy(src.alignment),
                'number_format': src.number_format,
                'fill': copy(src.fill),
            }
        ref_height = self.sheet.row_dimensions[ref_row].height

        # Insert one row per ER above stop_row (insert in reverse to preserve ordering)
        # After inserting N rows, stop_row shifts by N – track the insertion point
        insert_at = stop_row  # rows are inserted BEFORE this row
        for er in er_numbers:
            self.sheet.insert_rows(insert_at)

            # Apply reference row formatting to every cell in the new row
            for col_idx in range(1, max_col + 1):
                new_cell = self.sheet.cell(row=insert_at, column=col_idx)
                s = ref_styles[col_idx]
                new_cell.font = copy(s['font'])
                # Normalise top border: always use 'thin' so each new row looks
                # identical to the reference data rows regardless of insertion order
                orig_border = s['border']
                new_cell.border = Border(
                    left=copy(orig_border.left),
                    right=copy(orig_border.right),
                    top=Side(border_style='thin'),
                    bottom=copy(orig_border.bottom),
                )
                new_cell.alignment = copy(s['alignment'])
                new_cell.number_format = s['number_format']
                # Fill: green on PO column, no fill on all other columns
                if col_idx == po_col_idx:
                    new_cell.fill = green_fill
                else:
                    new_cell.fill = copy(s['fill'])

            if ref_height:
                self.sheet.row_dimensions[insert_at].height = ref_height

            # Write ER number into the PO column and WBS into column F
            self.sheet.cell(row=insert_at, column=po_col_idx, value=er)
            self.sheet.cell(row=insert_at, column=6, value=er_wbs_map.get(er, "ER"))
            # Write Vendor Name into column E
            vendor_val = er_vendor_map.get(er)
            if vendor_val is not None:
                self.sheet.cell(row=insert_at, column=5, value=vendor_val)
            # Write Cost Center into column J
            cc_val = er_cc_map.get(er)
            if cc_val is not None:
                self.sheet.cell(row=insert_at, column=10, value=cc_val)
            # Write G/L Account into column G
            gl_val = er_gl_map.get(er)
            if gl_val is not None:
                self.sheet.cell(row=insert_at, column=7, value=gl_val)
            # Write LE into column H
            le_val = er_le_map.get(er)
            if le_val is not None:
                self.sheet.cell(row=insert_at, column=8, value=le_val)
            # Write Country into column I
            country_val = er_country_map.get(er)
            if country_val is not None:
                self.sheet.cell(row=insert_at, column=9, value=country_val)
            # Write Gross PO Value into column K
            gross_val = er_gross_map.get(er)
            if gross_val is not None:
                self.sheet.cell(row=insert_at, column=11, value=round(gross_val, 2))
            # Write Requisition Title into column L
            req_val = er_req_map.get(er)
            if req_val is not None:
                self.sheet.cell(row=insert_at, column=12, value=req_val)

            # Write Total 2026 formula for this ER row
            self._write_total_formula(insert_at)

            pos[er] = insert_at
            insert_at += 1  # next ER goes after the one just inserted

        print(f"Inserted {len(er_numbers)} ER row(s) into template: {er_numbers}")
        return pos

    def _get_comments_col(self) -> str | None:
        """Scan the header row for a cell containing 'Comments'. Returns the column letter or None."""
        for col in range(1, (self.sheet.max_column or 200) + 1):
            val = self.sheet.cell(row=self.header_row, column=col).value
            if val and 'comment' in str(val).strip().lower():
                return get_column_letter(col)
        return None

    def write_hierarchy(self, hierarchy: dict, pos: dict[str, int]):
        '''
        Writes hierarchy to template.
        Iterates CostCenter -> WBSCode -> PO and writes MonthlyMetrics to correct cells.
        Only writes to blank cells unless overwrite=True.
        '''
        reclass_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        comments_col = self._get_comments_col()

        for cc_id, cost_center in hierarchy.items():
            for wbs_code, wbs in cost_center.wbs_codes.items():
                for po_number, po in wbs.pos.items():
                    # Normalize PO key before lookup so "9500905777.0" matches "9500905777"
                    po_number = self._norm_po(po_number)
                    # Skip if PO not in template
                    if po_number not in pos:
                        print(f"PO '{po_number}' not found in template. Skipping.")
                        continue
                    row = pos[po_number]
                    # Write Cost Center into column J if blank (or overwrite=True)
                    cc_cell = self.sheet.cell(row=row, column=10)
                    if self.overwrite or cc_cell.value is None or str(cc_cell.value).strip() == "":
                        cc_cell.value = cc_id
                    # Write Vendor Name into column E if blank/placeholder (or overwrite=True)
                    if po.vendor_name is not None:
                        vendor_cell = self.sheet.cell(row=row, column=5)
                        existing = str(vendor_cell.value).strip() if vendor_cell.value is not None else ""
                        if self.overwrite or existing == "" or existing.lower() == "not assigned":
                            vendor_cell.value = po.vendor_name
                    # Write WBS code into column F if blank (or overwrite=True)
                    # For ER rows use the real WBS from the transactional file if available
                    wbs_to_write = po.real_wbs if (wbs_code == "ER" and po.real_wbs) else wbs_code
                    wbs_cell = self.sheet.cell(row=row, column=6)
                    if self.overwrite or wbs_cell.value is None or str(wbs_cell.value).strip() == "":
                        wbs_cell.value = wbs_to_write
                    # Write G/L Account into column G if blank (or overwrite=True)
                    if po.gl_account is not None:
                        gl_cell = self.sheet.cell(row=row, column=7)
                        if self.overwrite or gl_cell.value is None or str(gl_cell.value).strip() == "":
                            gl_cell.value = po.gl_account
                    # Write LE (legal entity) into column H if blank (or overwrite=True)
                    if po.legal_entity is not None:
                        le_cell = self.sheet.cell(row=row, column=8)
                        if self.overwrite or le_cell.value is None or str(le_cell.value).strip() == "":
                            le_cell.value = po.legal_entity
                    # Write Country into column I if blank (or overwrite=True)
                    if po.country is not None:
                        country_cell = self.sheet.cell(row=row, column=9)
                        if self.overwrite or country_cell.value is None or str(country_cell.value).strip() == "":
                            country_cell.value = po.country
                    # Write Gross PO Value into column K if blank (or overwrite=True)
                    if po.gross_po_value is not None:
                        gross_cell = self.sheet.cell(row=row, column=11)
                        if self.overwrite or gross_cell.value is None or str(gross_cell.value).strip() == "":
                            gross_cell.value = round(po.gross_po_value, 2)
                    # Write Requisition Title into column L if blank (or overwrite=True)
                    if po.req_title is not None:
                        req_cell = self.sheet.cell(row=row, column=12)
                        if self.overwrite or req_cell.value is None or str(req_cell.value).strip() == "":
                            req_cell.value = po.req_title
                    for month, metrics in po.monthly_data.items():
                        if month not in self.column_map:
                            continue
                        month_cols = self.column_map[month]
                        values = {
                            'Accrual Reversal': metrics.accrual_reversal,
                            'Forecast': metrics.forecast,
                            'Accrual': metrics.accrual,
                            'Actual': metrics.actual
                        }
                        for metric, col_letter in month_cols.items():
                            cell = self.sheet[f"{col_letter}{row}"]
                            value = values.get(metric)
                            # Skip writing zero/None — avoids polluting blank template cells
                            # with 0 when there is genuinely no data for this metric/month.
                            # Always write when overwrite=True so existing values can be cleared.
                            if not self.overwrite and (value is None or value == 0 or value == 0.0):
                                continue
                            if self.overwrite or cell.value is None or str(cell.value).strip() == "":
                                cell.value = value

                        actual_col = month_cols.get('Actual')
                        accrual_col = month_cols.get('Accrual')
                        if actual_col and accrual_col:
                            variance_col = get_column_letter(column_index_from_string(actual_col) + 1)
                            variance_cell = self.sheet[f"{variance_col}{row}"]
                            if self.overwrite or variance_cell.value is None or str(variance_cell.value).strip() == "":
                                variance_cell.value = f"={actual_col}{row}-{accrual_col}{row}"

                        # Annotate Actual cell if reclass adjustments exist for this month
                        reclass_entries = po.reclass_adjustments.get(month, [])
                        if reclass_entries and actual_col:
                            actual_cell = self.sheet[f"{actual_col}{row}"]
                            # Highlight cell in light yellow
                            actual_cell.fill = reclass_fill
                            # Build comment text
                            lines = ["Reclass adjustment(s) included in Actual:"]
                            for amt, desc in reclass_entries:
                                sign = "+" if amt >= 0 else ""
                                lines.append(f"  {sign}${amt:,.2f}  —  {desc}")
                            note_text = "\n".join(lines)
                            # Add cell comment (tooltip)
                            comment = Comment(note_text, "Financial Automation")
                            comment.width = 400
                            comment.height = 120 + 20 * len(reclass_entries)
                            actual_cell.comment = comment
                            # Also write to Comments column if it exists
                            if comments_col:
                                comments_cell = self.sheet[f"{comments_col}{row}"]
                                existing = comments_cell.value or ""
                                separator = "\n" if existing else ""
                                comments_cell.value = existing + separator + note_text
                    # Write Total 2026 formula after all months are written for this PO
                    self._write_total_formula(row)

        # After all PO data is written, refresh the summary row formulas so they
        # cover the full data range (rows header_row+1 … last data row).
        self._update_summary_formulas()


    def _update_summary_formulas(self):
        """Rewrite range references in summary rows (rows 1 … header_row-1) so
        they span from header_row+1 to the actual last data row.

        The template ships with hardcoded ranges like K17:K18 (only 2 placeholder
        rows).  After PO rows are inserted those ranges need updating to cover all
        real data rows, otherwise SUBTOTAL / SUM / SUMIFS totals show wrong values.

        Strategy: for every formula cell above the header row, replace every
        occurrence of <COL><start_row>:<COL><end_row> where start_row equals
        header_row+1 with the correct end row.  Handles plain string formulas
        and ArrayFormula objects.
        """
        from openpyxl.worksheet.formula import ArrayFormula
        import re

        data_first = self.header_row + 1  # first real data row (e.g. 17)

        # Find the actual last data row (row before 'Previous Period Invoices')
        data_last = data_first
        for r in range(data_first, (self.sheet.max_row or 1000) + 1):
            val = self.sheet.cell(row=r, column=1).value
            if val is not None and str(val).strip() in (
                "Previous Period Invoices", "EXPENSE END"
            ):
                data_last = r - 1
                break
            data_last = r

        if data_last < data_first:
            return  # nothing to do (empty template)

        def _replace_range(formula_text: str) -> str:
            """Replace <COL><data_first>:<COL><old_end> with <COL><data_first>:<COL><data_last>."""
            def replacer(m):
                col1, row1, col2 = m.group(1), m.group(2), m.group(3)
                if int(row1) == data_first:
                    return f"{col1}{row1}:{col2}{data_last}"
                return m.group(0)
            return re.sub(r'([A-Z]+)(\d+):([A-Z]+)\d+', replacer, formula_text)

        updated = 0
        for row in range(1, self.header_row):
            for col in range(1, (self.sheet.max_column or 200) + 1):
                cell = self.sheet.cell(row=row, column=col)
                v = cell.value
                if v is None:
                    continue
                if isinstance(v, ArrayFormula):
                    new_text = _replace_range(v.text)
                    if new_text != v.text:
                        cell.value = ArrayFormula(v.ref, new_text)
                        updated += 1
                elif isinstance(v, str) and v.startswith('='):
                    new_formula = _replace_range(v)
                    if new_formula != v:
                        cell.value = new_formula
                        updated += 1

        if updated:
            print(f"Updated {updated} summary formula(s) to cover rows "
                  f"{data_first}:{data_last}.")


    ## Methods to write source sheets
    def write_forecast_source_sheet(self, forecast_df, pos: dict[str, int]):
        # Method to write forecast source sheet. 
        # Filter to template POs
        if self.forecast_po_col not in forecast_df.columns:
            raise KeyError(f"Expected {self.forecast_po_col} column not found in forecast dataframe.")

        forecast_df[self.forecast_po_col] = (
            forecast_df[self.forecast_po_col]
            .apply(lambda x: str(int(float(x))) if str(x).replace('.','',1).isdigit() else str(x))
        )
        filtered_df = forecast_df[forecast_df[self.forecast_po_col].isin(pos.keys())]

        visible_cols = [c for c in self.forecast_source_cols if c in filtered_df.columns]
        hidden_cols = [c for c in filtered_df.columns if c not in visible_cols]

        ordered_cols = visible_cols + hidden_cols
        source_df = filtered_df[ordered_cols]

        ws = self.wb.create_sheet("Forecast Source Data")

        for col_idx, col_name in enumerate(source_df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=col_name)

        for row_idx, (_, row) in enumerate(source_df.iterrows(), start=2):
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        data_start = 2
        data_end = ws.max_row
        total_row = data_end + 1
        ws.cell(row=total_row, column=1, value="PO Total")

        for col_idx, col_name in enumerate(source_df.columns, start=1):
            if col_name in visible_cols and col_name != self.forecast_po_col:
                letter = get_column_letter(col_idx)
                formula = f"=SUBTOTAL(9,{letter}{data_start}:{letter}{data_end})"
                ws.cell(row=total_row, column=col_idx, value=formula)

        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"

        for col_idx, col_name in enumerate(source_df.columns, start=1):
            letter = get_column_letter(col_idx)
            max_len = len(str(col_name))
            for row_idx in range(2, len(source_df) + 2):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[letter].width = max_len + 2

        if hidden_cols:
            start_idx = len(visible_cols) + 1
            end_idx = len(visible_cols) + len(hidden_cols)
            ws.column_dimensions.group(
                get_column_letter(start_idx),
                get_column_letter(end_idx),
                hidden=True
            )

    def write_transactional_source_sheet(self, transactions_df, pos: dict[str, int]):
        # Method to write transactional detail source sheet
        # Include rows where PO is in the template OR where Type is Reclass
        # (Reclass rows have no PO so they would otherwise be filtered out)
        if self.transactional_po_col not in transactions_df.columns:
            raise KeyError(f"Expected {self.transactional_po_col} column not found in transactional dataframe.")

        transactions_df[self.transactional_po_col] = transactions_df[self.transactional_po_col].astype(str)
        in_template = transactions_df[self.transactional_po_col].isin(pos.keys())
        is_reclass = transactions_df['Type'] == 'Reclass' if 'Type' in transactions_df.columns else False
        source_df = transactions_df[in_template | is_reclass]

        visible_cols = [c for c in self.transactional_source_cols if c in source_df.columns]
        hidden_cols = [c for c in source_df.columns if c not in visible_cols]

        final_cols = visible_cols + hidden_cols
        source_df = source_df[final_cols]

        ws = self.wb.create_sheet("Transactions Source Data")

        for col_idx, col_name in enumerate(source_df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=col_name)

        for row_idx, (_, row) in enumerate(source_df.iterrows(), start=2):
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"

        for col_idx, col_name in enumerate(source_df.columns, start=1):
            letter = get_column_letter(col_idx)
            max_len = len(str(col_name))
            for row_idx in range(2, len(source_df) + 2):
                value = ws.cell(row=row_idx, column=col_idx).value
                if value is not None:
                    max_len = max(max_len, len(str(value)))
            ws.column_dimensions[letter].width = max_len + 2

        if hidden_cols:
            start_idx = len(visible_cols) + 1
            end_idx = len(visible_cols) + len(hidden_cols)
            ws.column_dimensions.group(
                get_column_letter(start_idx),
                get_column_letter(end_idx),
                hidden=True
            )


    
    @staticmethod
    def _extract_er_number(po_value):
        """If po_value contains an ER number (e.g. 'ER97054 - ARCH & TECH - ...'), return only 'ER97054'.
        Otherwise return the original value unchanged."""
        if po_value and isinstance(po_value, str):
            match = re.match(r'^(ER\d+)', po_value.strip(), re.IGNORECASE)
            if match:
                return match.group(1).upper()
        return po_value

    def write_exception_sheet(self, exception_log, transactional_df):
        ws = self.wb.create_sheet("Exceptions")

        # Exclude reclasses from the exceptions tab
        class _FilteredLog:
            def __init__(self, entries):
                self.entries = entries
        exception_log = _FilteredLog([
            e for e in exception_log.entries
            if e.exception_type != ExceptionType.RECLASS
        ])

        # Define visible columns - include Description and GL Line Description from transactional data
        visible_headers = [
            'Cost Center', 'Accounting Period', 'WBS', 'PO/ER Number',
            'Exception Type', 'Source Row', 'Amount', 'Type', 'Description', 'GL Line Description'
        ]
        
        # Get all transactional columns for hidden section
        # Exclude columns already shown in visible section
        excluded_cols = {'Cost Center*', 'WBS Element', 'PO Number', 'Accounting Period', 'GL BER Corp Amount', 'Type', 'Description', 'GL Line Description'}
        hidden_headers = [col for col in transactional_df.columns if col not in excluded_cols]
        
        all_headers = visible_headers + hidden_headers
        
        # Write headers
        for col_idx, header in enumerate(all_headers, start=1):
            ws.cell(row=1, column=col_idx, value=header)
        
        # Write data rows
        for row_idx, entry in enumerate(exception_log.entries, start=2):
            # Visible columns
            ws.cell(row=row_idx, column=1, value=entry.cost_center)
            ws.cell(row=row_idx, column=2, value=entry.month)
            ws.cell(row=row_idx, column=3, value=entry.wbs)
            ws.cell(row=row_idx, column=4, value=self._extract_er_number(entry.po))
            ws.cell(row=row_idx, column=5, value=entry.exception_type.value)
            ws.cell(row=row_idx, column=6, value=entry.row_index)
            ws.cell(row=row_idx, column=7, value=entry.amount)
            ws.cell(row=row_idx, column=8, value=entry.transaction_type)
            # Description from source row data
            ws.cell(row=row_idx, column=9, value=entry.source_row_data.get('Description') if entry.source_row_data else None)
            # GL Line Description from source row data
            ws.cell(row=row_idx, column=10, value=entry.source_row_data.get('GL Line Description') if entry.source_row_data else None)
            
            # Hidden columns (remaining source row data)
            if entry.source_row_data:
                for col_idx_hidden, col_name in enumerate(hidden_headers, start=11):
                    ws.cell(row=row_idx, column=col_idx_hidden,
                           value=entry.source_row_data.get(col_name))
        
        # Apply formatting
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"
        
        # Auto-size visible columns
        for col_idx in range(1, len(visible_headers) + 1):
            letter = get_column_letter(col_idx)
            max_len = len(str(all_headers[col_idx - 1]))
            for row_idx in range(2, len(exception_log.entries) + 2):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None:
                    max_len = max(max_len, len(str(cell_value)))
            ws.column_dimensions[letter].width = min(max_len + 2, 50)
        
        # Group and hide supplementary columns
        if hidden_headers:
            start_idx = len(visible_headers) + 1
            end_idx = len(all_headers)
            ws.column_dimensions.group(
                get_column_letter(start_idx),
                get_column_letter(end_idx),
                hidden=True
            )

    def write_exception_data_sheet(self, exception_log):
        """Write raw exception data to hidden sheet for formula reference"""
        ws = self.wb.create_sheet("Exception_Data")

        # Exclude reclasses
        class _FilteredLog:
            def __init__(self, entries):
                self.entries = entries
        exception_log = _FilteredLog([
            e for e in exception_log.entries
            if e.exception_type != ExceptionType.RECLASS
        ])

        # Headers
        headers = ['Cost Center', 'WBS', 'PO', 'Exception Type', 'Accounting Period', 'Amount', 'Type']
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_idx, value=header)
            ws.cell(row=1, column=col_idx).font = Font(bold=True)
        
        # Data rows
        for row_idx, entry in enumerate(exception_log.entries, start=2):
            ws.cell(row=row_idx, column=1, value=entry.cost_center or '')
            ws.cell(row=row_idx, column=2, value=entry.wbs or '')
            ws.cell(row=row_idx, column=3, value=self._extract_er_number(entry.po) or '')
            ws.cell(row=row_idx, column=4, value=entry.exception_type.value)
            ws.cell(row=row_idx, column=5, value=entry.month or '')
            ws.cell(row=row_idx, column=6, value=entry.amount)
            ws.cell(row=row_idx, column=7, value=entry.transaction_type or '')
        
        # Hide the sheet
        ws.sheet_state = 'hidden'
    
    def write_exception_summary_sheet(self, exception_log):
        """Create a summary sheet with interactive month filter showing exception counts by type and by cost center"""
        ws = self.wb.create_sheet("Exceptions Summary")

        # Exclude reclasses — build a filtered proxy that exposes only the methods used below
        filtered_entries = [
            e for e in exception_log.entries
            if e.exception_type != ExceptionType.RECLASS
        ]
        from collections import Counter
        class _FilteredLog:
            def __init__(self, entries):
                self.entries = entries
            def summary_by_type(self):
                counts = Counter(e.exception_type.value for e in self.entries)
                total = len(self.entries)
                return {'counts': dict(counts), 'total': total,
                        'percentages': {k: (v/total*100) if total > 0 else 0 for k, v in counts.items()}}
            def summary_by_cost_center(self):
                result = {}
                for entry in self.entries:
                    cc = entry.cost_center or 'Unknown'
                    exc_type = entry.exception_type.value
                    if cc not in result:
                        result[cc] = {'total': 0, 'by_type': {}}
                    result[cc]['total'] += 1
                    result[cc]['by_type'][exc_type] = result[cc]['by_type'].get(exc_type, 0) + 1
                return result
        exception_log = _FilteredLog(filtered_entries)

        # Get summary data for getting unique values
        summary_by_type = exception_log.summary_by_type()
        summary_by_cc = exception_log.summary_by_cost_center()
        
        # Get all unique exception types and cost centers
        all_exception_types = sorted(set(summary_by_type['counts'].keys()))
        all_cost_centers = sorted(set(summary_by_cc.keys()))
        
        # Get all unique months from exception data — normalise to 3-letter names where possible
        _month_norm = {
            'january':'Jan','february':'Feb','march':'Mar','april':'Apr',
            'may':'May','june':'Jun','july':'Jul','august':'Aug',
            'september':'Sep','october':'Oct','november':'Nov','december':'Dec'
        }
        def _norm_month(m):
            s = str(m).strip()
            return _month_norm.get(s.lower(), s)

        all_months = sorted(
            set(_norm_month(entry.month) for entry in exception_log.entries if entry.month),
            key=month_sort_key
        )

        current_row = 1

        # Add Month Filter Dropdown
        ws.cell(row=current_row, column=1, value="Filter by Month:")
        ws.cell(row=current_row, column=1).font = Font(bold=True, size=12)

        # Create dropdown list — Excel DataValidation formula1 has a 255-char limit.
        # Write months to a helper column (ZZ) and reference that range instead.
        month_options = ["All Months"] + all_months
        helper_col = 100  # column CV — well out of view
        for i, opt in enumerate(month_options, start=1):
            ws.cell(row=i, column=helper_col, value=opt)
        helper_letter = get_column_letter(helper_col)
        helper_range = f"'{ws.title}'!${helper_letter}$1:${helper_letter}${len(month_options)}"
        dv = DataValidation(type="list", formula1=helper_range, allow_blank=False)
        dv.add(ws['B1'])
        ws.add_data_validation(dv)

        # Set default value
        ws['B1'] = "All Months"
        ws['B1'].font = Font(size=11)
        
        # Define named range for the filter cell
        from openpyxl.workbook.defined_name import DefinedName
        defined_name = DefinedName('MonthFilter', attr_text=f"'{ws.title}'!$B$1")
        self.wb.defined_names['MonthFilter'] = defined_name
        
        current_row += 2
        
        # Section 1: Summary by Exception Type (with dynamic formulas)
        ws.cell(row=current_row, column=1, value="Exceptions Summary by Type")
        ws.cell(row=current_row, column=1).font = Font(bold=True, size=14)
        current_row += 2
        
        # Headers for type summary
        ws.cell(row=current_row, column=1, value="Exception Type")
        ws.cell(row=current_row, column=2, value="Count")
        ws.cell(row=current_row, column=3, value="% of Total")
        for col in range(1, 4):
            ws.cell(row=current_row, column=col).font = Font(bold=True)
        current_row += 1
        
        # Data rows for type summary with formulas
        type_start_row = current_row
        for exc_type in all_exception_types:
            ws.cell(row=current_row, column=1, value=exc_type)
            
            # Count formula: IF MonthFilter="All Months", count all, else count for specific month
            count_formula = (
                f'=IF(MonthFilter="All Months",'
                f'COUNTIF(Exception_Data!$D:$D,"{exc_type}"),'
                f'COUNTIFS(Exception_Data!$D:$D,"{exc_type}",Exception_Data!$E:$E,MonthFilter))'
            )
            ws.cell(row=current_row, column=2, value=count_formula)
            
            # Percentage formula
            pct_formula = f'=IF(SUM($B${type_start_row}:$B${type_start_row + len(all_exception_types) - 1})=0,0,B{current_row}/SUM($B${type_start_row}:$B${type_start_row + len(all_exception_types) - 1})*100)'
            ws.cell(row=current_row, column=3, value=pct_formula)
            ws.cell(row=current_row, column=3).number_format = '0.0"%"'
            
            current_row += 1
        
        # Total row for type summary
        ws.cell(row=current_row, column=1, value="TOTAL")
        ws.cell(row=current_row, column=1).font = Font(bold=True)
        total_formula = f'=SUM(B{type_start_row}:B{current_row - 1})'
        ws.cell(row=current_row, column=2, value=total_formula)
        ws.cell(row=current_row, column=2).font = Font(bold=True)
        ws.cell(row=current_row, column=3, value='100.0%')
        ws.cell(row=current_row, column=3).font = Font(bold=True)
        current_row += 3
        
        # Section 2: Summary by Cost Center (with dynamic formulas)
        ws.cell(row=current_row, column=1, value="Exceptions Summary by Cost Center")
        ws.cell(row=current_row, column=1).font = Font(bold=True, size=14)
        current_row += 2
        
        # Headers for cost center summary
        ws.cell(row=current_row, column=1, value="Cost Center")
        ws.cell(row=current_row, column=2, value="Total")
        for idx, exc_type in enumerate(all_exception_types, start=3):
            ws.cell(row=current_row, column=idx, value=exc_type)
        for col in range(1, len(all_exception_types) + 3):
            ws.cell(row=current_row, column=col).font = Font(bold=True)
        current_row += 1
        
        # Data rows for cost center summary with formulas
        cc_start_row = current_row
        for cc in all_cost_centers:
            ws.cell(row=current_row, column=1, value=cc)
            
            # Total formula for this cost center
            total_formula = (
                f'=IF(MonthFilter="All Months",'
                f'COUNTIF(Exception_Data!$A:$A,"{cc}"),'
                f'COUNTIFS(Exception_Data!$A:$A,"{cc}",Exception_Data!$E:$E,MonthFilter))'
            )
            ws.cell(row=current_row, column=2, value=total_formula)
            
            # Count by exception type
            for idx, exc_type in enumerate(all_exception_types, start=3):
                type_formula = (
                    f'=IF(MonthFilter="All Months",'
                    f'COUNTIFS(Exception_Data!$A:$A,"{cc}",Exception_Data!$D:$D,"{exc_type}"),'
                    f'COUNTIFS(Exception_Data!$A:$A,"{cc}",Exception_Data!$D:$D,"{exc_type}",Exception_Data!$E:$E,MonthFilter))'
                )
                ws.cell(row=current_row, column=idx, value=type_formula)
            
            current_row += 1
        
        # Auto-size columns
        max_cols = max(len(all_exception_types) + 3, 10)
        for col_idx in range(1, max_cols):
            letter = get_column_letter(col_idx)
            max_len = 10
            for row_idx in range(1, current_row):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None and not str(cell_value).startswith('='):
                    max_len = max(max_len, len(str(cell_value)))
            ws.column_dimensions[letter].width = min(max_len + 2, 50)
        
        # Freeze panes at row 4 (below filter and title)
        ws.freeze_panes = "A4"
        
    
    def save(self):
        """Saves the workbook to the output path."""
        try:
            self.wb.save(self.output_path)
            print(f"Workbook saved to: {self.output_path}")
        except Exception as e:
            raise Exception(f"Failed to save workbook: {e}")
        
