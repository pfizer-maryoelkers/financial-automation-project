
import pandas as pd
import re

class TransactionalDetailReader:
    """Reads a C-TIES or Consolidated Actuals file and extracts typed, month-bucketed transactions.

    Classification rules (in priority order):
        1. Document number / Vendor Invoice / AP Voucher Number prefix + amount sign:
               "2" + positive → Accrual,  "2" + negative → Reversal
               "5" → Actual
               "9" → Reclass
        2. CO Doc Line Item Txt keywords (fallback when no prefix matches):
               reversal → Reversal, accrual → Accrual, reclass → Reclass, invoice/vendor → Actual
        3. Amount sign fallback for unresolved rows

    Month mapping is direct: AP01 = Jan, AP02 = Feb, … AP12 = Dec.
    Two Decembers are distinguished by year: the smaller year → "Dec (PY)".
    """

    def __init__(self, file_path, required_cols, valid_types, colmap):
        """Initialize with the transactional detail file path.
        
        See config_base.yaml for default parameters
        """
        self.file_path = file_path
        self.data = None
        self._excel_row: 'pd.Series | None' = None  # maps pandas index → 1-based Excel row

        # Strip whitespace from required_cols and colmap values so they always
        # match file headers regardless of accidental leading/trailing spaces in config.
        self.required_cols = {c.strip() for c in required_cols}

        self.valid_types = valid_types # Valid types for reading (currently support Actuals, Accruals, Reversals) - open to extension

        self.colmap = {k: v.strip() if isinstance(v, str) else v for k, v in colmap.items()}

        self.month_map = {
            1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
            7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
        } # Month map for reading



    @staticmethod
    def _find_po_column(columns) -> str | None:
        """Find the PO number column by scanning column names with loose formatting."""
        cols = [str(c).strip() for c in columns]
        
        # Priority 1: Exact or highly specific matches
        for c in cols:
            c_upper = c.upper()
            if c_upper in ('PO', 'PO#', 'PO NUMBER', 'GL PO NUMBER', 'PURCHASE ORDER', 'PO_NO', 'PO NO', 'PO_NUMBER'):
                return c
                
        # Priority 2: Contains 'PO#' or 'PO_NO' or '/ PO#' or '/PO#' or 'num/PO#'
        for c in cols:
            c_upper = c.upper()
            if any(x in c_upper for x in ('DESC', 'DATE', 'AMT', 'AMOUNT', 'VALUE', 'TITLE', 'STATUS', 'ITEM', 'PRICE', 'COST')):
                continue
            if 'PO#' in c_upper or 'PO_NO' in c_upper or 'PO NO' in c_upper or 'NUM/PO#' in c_upper or '/ PO#' in c_upper or '/PO#' in c_upper:
                return c
                
        # Priority 3: Contains 'PO' as a word boundary, e.g. "GL PO" or "Document Number / PO"
        for c in cols:
            c_upper = c.upper()
            if any(x in c_upper for x in ('DESC', 'DATE', 'AMT', 'AMOUNT', 'VALUE', 'TITLE', 'STATUS', 'ITEM', 'PRICE', 'COST')):
                continue
            if re.search(r'\bPO\b', c_upper) or 'PO/' in c_upper or '/PO' in c_upper or '/ PO' in c_upper:
                return c
                
        return None


    @staticmethod
    def _strip_col_headers(df: 'pd.DataFrame') -> 'pd.DataFrame':
        """Strip leading/trailing whitespace from all column names and
        collapse any runs of internal whitespace to a single space.
        This handles column names like 'WBS/Internal Order ' (trailing space)
        or 'Cost  Center' (double space) that appear in some source files."""
        df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)
        return df

    def _parse_ap(self, raw_month) -> 'tuple[int, int] | tuple[None, None]':
        """Parse a raw Accounting Period value to a (month_num, year) tuple.

        AP values seen in practice:
          • "AP202501"  / 202501  → (1, 2025)   i.e. January 2025
          • "AP202512"  / 202512  → (12, 2025)  i.e. December 2025  ← prior-year Dec
          • "AP202612"  / 202612  → (12, 2026)  i.e. December 2026  ← current-year Dec
          • "Period 01 2026"      → (1, 2026)
          • plain integer 1–12   → (month, None)  — year unknown, treated as current year
          • pandas Timestamp      → (month, year)

        Returns (None, None) when the value cannot be parsed or month is out of range.
        """
        import pandas as _pd
        if isinstance(raw_month, _pd.Timestamp):
            return raw_month.month, raw_month.year

        try:
            raw_str = str(raw_month).strip()

            # "Period 01 2026" / "Period 01"
            period_match = re.search(r'period\s+(\d+)(?:\s+(\d{4}))?', raw_str, re.IGNORECASE)
            if period_match:
                month_num = int(period_match.group(1))
                year = int(period_match.group(2)) if period_match.group(2) else None
                return (month_num, year) if 1 <= month_num <= 12 else (None, None)

            # Strip leading non-digit characters (e.g. "AP" prefix)
            digits_only = re.sub(r'^[^0-9]+', '', raw_str)
            if not digits_only:
                return None, None

            raw_int = int(float(digits_only))
            if raw_int > 12:
                # YYYYMM format: e.g. 202501 → month=1, year=2025
                year = raw_int // 100
                month_num = raw_int % 100
            else:
                # Plain 1–12 — year unknown
                month_num = raw_int
                year = None

        except (TypeError, ValueError):
            return None, None

        return (month_num, year) if 1 <= month_num <= 12 else (None, None)

    def _parse_month_num(self, raw_month) -> 'int | None':
        """Convenience wrapper around _parse_ap that returns only the month number.
        Used by callers that do not need year-awareness (reclass helpers, etc.)."""
        month, _ = self._parse_ap(raw_month)
        return month

    def _detect_header_row(self, sheet_name: str) -> int:
        """Return the 0-based header row index for a sheet.

        Scans the first 10 rows (indices 0-9).  For each candidate row the
        sheet is read with that row as the header; the column names are stripped
        and checked against the required-column alias list.  Returns the first
        matching row index.

        Skips any candidate row whose columns are all numeric (summary/totals
        rows that appear above the real header in many Consolidated files).

        Falls back to 0 when nothing matches — the caller surfaces the error.
        """
        for header in range(10):
            try:
                preview = pd.read_excel(
                    self.file_path, sheet_name=sheet_name,
                    header=header, nrows=5
                )
            except Exception:
                break
            preview = self._strip_col_headers(preview)
            # Skip rows whose column names are all numeric — those are totals rows
            non_numeric_cols = [
                c for c in preview.columns
                if not str(c).strip().lstrip('-').replace('.', '', 1).isdigit()
            ]
            if not non_numeric_cols:
                continue
            if self._sheet_has_required_cols(preview):
                return header
        return 0  # fallback

    def _normalise_cols(self, df: 'pd.DataFrame') -> 'pd.DataFrame':
        """Rename format-specific column names to the standard names used by all
        downstream logic.  Handles two Consolidated Actuals variants and C-TIES:

        • C-TIES                 — standard format, minimal renames needed
        • Consolidated AP07 2026 — WBS HIERARCHY, Document Number / PO#,
                                   Fiscal Year/Period, Vendor Desc, Company code
        • Consolidated Shalisha  — WBS/Internal Order, Document num/PO#,
                                   Period, Vendor Description, Company Code
        Both Consolidated variants carry a 'P3 ID' column which is the reliable
        cost-center identifier — mapped to the cost_center target so downstream
        P3 ID matching works correctly (Cost Center is '#' / not assigned in both).
        """
        rename: dict[str, str] = {}

        # ── Period / Month ───────────────────────────────────────────────────
        month_target = self.colmap.get('month', 'Accounting Period')
        if month_target not in df.columns:
            for alias in ('Accounting Period', 'Fiscal Year/Period', 'Period', 'Month'):
                if alias in df.columns:
                    rename[alias] = month_target
                    break

        # ── PO Number ────────────────────────────────────────────────────────
        po_target  = self.colmap.get('po', 'PO Number')
        cls_target = self.colmap.get('classifier', 'AP Voucher Number')
        if po_target not in df.columns:
            found_po = self._find_po_column(df.columns)
            if found_po:
                # When the Consolidated file also has a Vendor Invoice column,
                # that column is the real classifier (2=Accrual/Reversal,
                # 5=Invoice, 9=Reclass).  Do NOT copy the PO column into
                # cls_target — leave cls_target to be filled by Vendor Invoice.
                if (
                    cls_target != po_target
                    and cls_target not in df.columns
                    and 'Vendor Invoice' not in df.columns
                ):
                    df = df.copy()
                    df[cls_target] = df[found_po]
                rename[found_po] = po_target

        # ── WBS ──────────────────────────────────────────────────────────────
        wbs_target = self.colmap.get('wbs', 'WBS Element')
        if wbs_target not in df.columns:
            for alias in (
                'Planful WBS Element',
                'WBS HIERARCHY',        # AP07 2026
                'WBS Hierarchy',        # Shalisha AP03
                'WBS Element',
                'WBS/Internal Order',   # Shalisha AP12 (trailing space normalised by _strip_col_headers)
                'WBS',
            ):
                if alias in df.columns:
                    rename[alias] = wbs_target
                    break

        # ── Amount (BER) ─────────────────────────────────────────────────────
        amt_target = self.colmap.get('amount', 'GL BER Corp Amount')
        
        # Priority mapping for Amount columns:
        # We prefer columns containing 'BER' (such as 'Amount - BER' or 'GL BER Corp Amount') over generic 'Amount'
        best_amt_col = None
        cols_set = set(df.columns)
        
        if 'Amount - BER' in cols_set:
            best_amt_col = 'Amount - BER'
        else:
            ber_cols = [c for c in df.columns if 'BER' in c]
            if ber_cols:
                if amt_target in ber_cols:
                    best_amt_col = amt_target
                else:
                    best_amt_col = ber_cols[0]
            elif amt_target in cols_set:
                best_amt_col = amt_target
            elif 'Amount - MAR' in cols_set:
                best_amt_col = 'Amount - MAR'
            elif 'Amount' in cols_set:
                best_amt_col = 'Amount'
                
        if best_amt_col and best_amt_col != amt_target:
            if amt_target in df.columns:
                df = df.copy()
                df = df.drop(columns=[amt_target])
            rename[best_amt_col] = amt_target

        # ── Cost Center / P3 ID ──────────────────────────────────────────────
        # Both Consolidated formats have a 'P3 ID' column that holds the real
        # project identifier.  'Cost Center' is '#' / not assigned in both files.
        # Map 'P3 ID' → cc_target so downstream P3-ID matching works correctly.
        cc_target = self.colmap.get('cost_center', 'Cost Center*')
        if cc_target not in df.columns:
            for alias in ('P3 ID', 'Cost Center', 'CC ID', 'Cost Center*'):
                if alias in df.columns:
                    rename[alias] = cc_target
                    break
        elif 'P3 ID' in df.columns and cc_target in df.columns:
            # cc_target exists but may be '#' — overwrite with the real P3 ID value
            # where cc_target is blank/placeholder and P3 ID is populated.
            _placeholder = {'#', '', 'nan', 'none', 'not assigned', 'wapc/not assigned'}
            _cc_blank = df[cc_target].astype(str).str.strip().str.lower().isin(_placeholder)
            _p3_valid = df['P3 ID'].notna() & ~df['P3 ID'].astype(str).str.strip().str.lower().isin(_placeholder)
            df = df.copy()
            df.loc[_cc_blank & _p3_valid, cc_target] = df.loc[_cc_blank & _p3_valid, 'P3 ID']

        # ── Legal Entity ─────────────────────────────────────────────────────
        le_target = self.colmap.get('legal_entity', 'Legal Entity')
        if le_target not in df.columns:
            for alias in (
                'Company code',   # AP07 2026
                'Company Code',   # Shalisha AP12
            ):
                if alias in df.columns:
                    rename[alias] = le_target
                    break

        # ── Vendor Name ──────────────────────────────────────────────────────
        vendor_target = self.colmap.get('vendor_name', 'Vendor Name')
        if vendor_target not in df.columns:
            for alias in (
                'Vendor Name',          # direct match when target differs
                'Vendor Desc',          # AP07 2026 Consolidated
                'Vendor Description',   # Shalisha AP12
            ):
                if alias in df.columns:
                    rename[alias] = vendor_target
                    break

        # ── Project Name ─────────────────────────────────────────────────────
        project_name_target = self.colmap.get('project_name', 'Project Name')
        if project_name_target not in df.columns:
            for alias in (
                'Project Name',         # direct match when target differs
                'Project Description',
                'Project',
            ):
                if alias in df.columns:
                    rename[alias] = project_name_target
                    break

        # ── Classifier (AP Voucher Number) ───────────────────────────────────
        # Vendor Invoice is the real classifier for the Consolidated format
        # (2=Accrual/Reversal, 5=Invoice/Actual, 9=Reclass).  When it exists,
        # always prefer it over anything already in cls_target (which may have
        # been filled from the PO column above).
        if 'Vendor Invoice' in df.columns:
            if 'Vendor Invoice' not in rename.values():  # not already being renamed
                df = df.copy()
                df[cls_target] = df['Vendor Invoice']
        elif cls_target not in df.columns:
            for alias in ('AP Voucher Number',):
                if alias in df.columns:
                    rename[alias] = cls_target
                    break

        # Apply all renames at once
        if rename:
            df = df.rename(columns=rename)

        # ── Synthetic Type column — Consolidated has no Type column;
        #    every row in that file is an actual transaction.
        type_col = self.colmap.get('type', 'Type')
        if type_col not in df.columns:
            df[type_col] = 'Actual'

        return df

    def _sheet_has_required_cols(self, preview: 'pd.DataFrame') -> bool:
        """Check required cols, accepting all known column-name aliases across
        the three supported source formats (C-TIES, Chargeout, Consolidated)."""
        cols = set(preview.columns)

        def _swap(aliases: list[str], target: str) -> None:
            """Replace the first matching alias in cols with target."""
            if target not in cols:
                for alias in aliases:
                    if alias in cols:
                        cols.discard(alias)
                        cols.add(target)
                        break

        # Period
        _swap(['Month', 'Fiscal Year/Period', 'Period', 'Accounting Period'], self.colmap.get('month', 'Accounting Period'))
        # PO Number — both Consolidated variants and loose formatting
        found_po_col = self._find_po_column(preview.columns)
        if found_po_col:
            _swap([found_po_col], self.colmap.get('po', 'PO Number'))
        else:
            _swap(['GL PO Number', 'Document Number / PO#', 'Document num/PO#', 'PO Number', 'PO#', 'PO'],
                  self.colmap.get('po', 'PO Number'))
        # WBS — all variants (_strip_col_headers normalises trailing/extra spaces)
        _swap(['Planful WBS Element', 'WBS HIERARCHY', 'WBS Hierarchy', 'WBS Element', 'WBS/Internal Order', 'WBS'],
              self.colmap.get('wbs', 'WBS Element'))
        
        # Amount - BER vs Amount
        ber_cols = [c for c in cols if isinstance(c, str) and 'BER' in c]
        if ber_cols:
            _swap(ber_cols + ['Amount - BER', 'Amount - MAR', 'Amount'], self.colmap.get('amount', 'GL BER Corp Amount'))
        else:
            _swap(['Amount - BER', 'Amount - MAR', 'Amount'], self.colmap.get('amount', 'GL BER Corp Amount'))

        _swap(['GL Transaction Amount', 'Amount - BER', 'Amount'], 'GL Transaction Amount')
        _swap(['Vendor Invoice', 'AP Voucher Number', 'Document Number / PO#', 'Document num/PO#'],
              self.colmap.get('classifier', 'AP Voucher Number'))

        # Treat any required col that matches a normalised alias as satisfied.
        # Also ensure Type is considered present — Consolidated synthesises it.
        type_col = self.colmap.get('type', 'Type')
        cols.add(type_col)
        # Mirror any required-col aliases so the subset check passes
        for req in list(self.required_cols):
            if req not in cols:
                for alias_list, target in [
                    (['Month', 'Fiscal Year/Period', 'Period', 'Accounting Period'], self.colmap.get('month', 'Accounting Period')),
                    ([found_po_col] if found_po_col else ['GL PO Number', 'Document Number / PO#', 'Document num/PO#', 'PO Number', 'PO#', 'PO'],
                     self.colmap.get('po', 'PO Number')),
                    (['Amount - BER', 'Amount - MAR', 'Amount'] + [c for c in preview.columns if isinstance(c, str) and 'BER' in c], self.colmap.get('amount', 'GL BER Corp Amount')),
                    (['GL Transaction Amount', 'Amount'], 'GL Transaction Amount'),
                ]:
                    if req in alias_list and target in cols:
                        cols.add(req)
                        break

        return self.required_cols.issubset(cols)

    def load_transactional_detail_file(self):
        """Load all valid sheets from the transactional Excel file.
        Valid sheet defined by having the following columns: 'PO Number', 'Accounting Period', 'GL Transaction Amount'
        Header row is auto-detected (row 1 or row 2). Column header whitespace is stripped automatically.
        Files that use 'Month' instead of 'Accounting Period' are accepted and normalised automatically.
        """
        try:
            xls = pd.ExcelFile(self.file_path)

            valid_sheets = []
            header_map: dict[str, int] = {}
            for sheet in xls.sheet_names:
                header = self._detect_header_row(sheet)
                preview = pd.read_excel(self.file_path, sheet_name=sheet, header=header, nrows=5)
                preview = self._strip_col_headers(preview)
                if self._sheet_has_required_cols(preview):
                    valid_sheets.append(sheet)
                    header_map[sheet] = header

            if not valid_sheets:
                # Compile a helpful diagnostic of why the sheets failed the required columns check
                sheet_diagnostics = []
                for sheet in xls.sheet_names:
                    try:
                        hdr = self._detect_header_row(sheet)
                        preview = pd.read_excel(self.file_path, sheet_name=sheet, header=hdr, nrows=5)
                        preview = self._strip_col_headers(preview)
                        # Find which required columns were missing
                        cols = set(preview.columns)
                        found_period = any(x in cols for x in ('Month', 'Fiscal Year/Period', 'Period', 'Accounting Period'))
                        found_po = self._find_po_column(preview.columns) is not None
                        found_wbs = any(x in cols for x in ('Planful WBS Element', 'WBS HIERARCHY', 'WBS Hierarchy', 'WBS Element', 'WBS/Internal Order', 'WBS'))
                        found_amount = any(isinstance(c, str) and 'BER' in c for c in cols) or any(x in cols for x in ('Amount - BER', 'Amount - MAR', 'Amount'))
                        
                        missing_status = []
                        if not found_period: missing_status.append("Period/Month")
                        if not found_po: missing_status.append("PO Number")
                        if not found_wbs: missing_status.append("WBS Code")
                        if not found_amount: missing_status.append("Amount")
                        
                        if missing_status:
                            sheet_diagnostics.append(f"Sheet '{sheet}': Missing {', '.join(missing_status)}")
                        else:
                            sheet_diagnostics.append(f"Sheet '{sheet}': Layout looks valid but other required check failed")
                    except Exception as sheet_err:
                        sheet_diagnostics.append(f"Sheet '{sheet}': Error scanning layout ({sheet_err})")
                
                raise ValueError(
                    f"No valid sheets found containing the required transactional columns in the file '{self.file_path}'. "
                    f"Required columns include PO Number, WBS Element, Period/Month, and Amount. "
                    f"All scanned worksheets in this workbook: {xls.sheet_names}. "
                    f"Detailed diagnostics per worksheet:\n  " + "\n  ".join(sheet_diagnostics)
                )

            print(f"Loading valid sheets: {valid_sheets}")

            raw_dfs = [
                self._normalise_cols(
                    self._strip_col_headers(
                        pd.read_excel(self.file_path, sheet_name=sheet, header=header_map[sheet])
                    )
                )
                for sheet in valid_sheets
            ]

            # Build a per-row Excel row number Series (1-based) before resetting the
            # index.  For each sheet, the first data row sits at Excel row
            # (header_map[sheet] + 2) — one for the header row itself, one for the
            # 1-based offset — and increments by 1 for each subsequent row.
            excel_row_arrays = []
            for sheet, df in zip(valid_sheets, raw_dfs):
                header_offset = header_map[sheet]  # 0-based header row in the Excel file
                start_excel_row = header_offset + 2  # first data row in 1-based Excel numbering
                excel_row_arrays.extend(range(start_excel_row, start_excel_row + len(df)))
            # Will be re-indexed after ignore_index concat (index 0..N-1 matches position)
            self._excel_row = pd.Series(
                excel_row_arrays,
                index=range(len(excel_row_arrays)),
                name='_excel_row',
                dtype='Int64',
            )

            self.data = pd.concat(raw_dfs, ignore_index=True)

            type_col = self.colmap.get('type', 'Type')
            # Detect whether the source already supplied a real Type column.
            # _normalise_cols fills a synthetic all-'Actual' column when the
            # source has no Type column at all (both Consolidated variants).
            # That synthetic fill does NOT count as "supplied" — we re-classify.
            type_was_present = type_col in self.data.columns and not (
                (self.data[type_col] == 'Actual').all()
                and 'AP Voucher Number' in self.data.columns
            )
            if not type_was_present:
                self.data[type_col] = self._categorize()

            # ── Normalise PO column (vectorised) ────────────────────────────────
            # Convert float-formatted integers: 9500905777.0 → "9500905777".
            # Strategy: try to cast the whole column to Int64 first (fast), then
            # fix up anything that failed individually.
            po_col_name = self.colmap['po']
            _po_s = self.data[po_col_name].astype(str).str.strip()
            # Save the original raw PO column values before any cleaning/normalization
            self.data['Original_PO_Doc_No'] = _po_s
            # Rows that look like plain integers (digits only, optionally .0)
            _int_mask = _po_s.str.fullmatch(r'-?\d+(?:\.\d+)?')
            if _int_mask.any():
                _po_s = _po_s.copy()
                try:
                    _po_s[_int_mask] = (
                        pd.to_numeric(_po_s[_int_mask], errors='coerce')
                        .dropna()
                        .astype('int64')
                        .astype(str)
                    )
                except (ValueError, OverflowError):
                    pass
            self.data[po_col_name] = _po_s

            # ── Convert Reclass rows without 95 POs into POs ──────────────────────
            # "If the P3 ID has a reclass and there is not a 9500863325 (Example) in
            #  the Document num column or in the CO Doc Line Item Txt line than you
            #  can treat it as a PO and put it on the front of the template.
            #  next to the 900 whatever please put RC so people know it was a reclass"
            type_col_name = self.colmap.get('type', 'Type')
            _co_doc_col   = 'CO Doc Line Item Txt'
            if type_col_name in self.data.columns and po_col_name in self.data.columns:
                is_reclass = self.data[type_col_name] == 'Reclass'
                if is_reclass.any():
                    po_series = self.data[po_col_name].astype(str).str.strip()
                    has_95_in_po = po_series.str.contains(r'\b95\d{8,}\b', regex=True, na=False)

                    has_95_in_co_doc = pd.Series(False, index=self.data.index)
                    if _co_doc_col in self.data.columns:
                        co_doc_series = self.data[_co_doc_col].astype(str).str.strip()
                        has_95_in_co_doc = co_doc_series.str.contains(r'\b95\d{8,}\b', regex=True, na=False)

                    reclass_to_po_mask = is_reclass & ~has_95_in_po & ~has_95_in_co_doc

                    # Ensure we have a valid non-empty document number
                    _PLACEHOLDERS = {'', 'none', 'nan', '#', 'not assigned'}
                    valid_doc_num = (
                        po_series.notna()
                        & (po_series != '')
                        & (~po_series.str.lower().isin(_PLACEHOLDERS))
                    )
                    final_mask = reclass_to_po_mask & valid_doc_num

                    if final_mask.any():
                        self.data.loc[final_mask, po_col_name] = po_series[final_mask] + " RC"
                        self.data.loc[final_mask, type_col_name] = 'Actual'
                        n_converted = int(final_mask.sum())
                        print(f"  - Converted {n_converted} Reclass row(s) without 95 PO to Actual PO rows (appended ' RC').")


            # ── Resolve non-95 PO rows via CO Doc Line Item Txt (vectorised) ────
            # Rows where Document num/PO# does NOT start with '95' are not real
            # vendor POs.  The actual PO may be embedded in CO Doc Line Item Txt:
            #   "AC01-- 9500852590 -AMPS ACCELERATION AP01 ACCRUALS"
            # Vectorised: use Series.str.extract (regex engine stays in C).
            if po_col_name in self.data.columns:
                po_series  = self.data[po_col_name].astype(str).str.strip()
                is_real_po = po_series.str.startswith('95')
                is_er_row  = (
                    self.data[type_col_name] == 'ER'
                    if type_col_name in self.data.columns
                    else pd.Series(False, index=self.data.index)
                )
                non_po_mask = ~is_real_po & ~is_er_row & ~po_series.str.endswith(' RC', na=False)

                if non_po_mask.any() and _co_doc_col in self.data.columns:
                    # str.extract returns the first capture group (vectorised C loop)
                    extracted = (
                        self.data.loc[non_po_mask, _co_doc_col]
                        .astype(str)
                        .str.extract(r'\b(95\d{8,})\b', expand=False)
                    )
                    resolved   = extracted.notna()
                    # Build full-length unresolved mask without dtype mismatch
                    unresolved = non_po_mask & ~resolved.reindex(self.data.index, fill_value=False)

                    self.data.loc[non_po_mask & resolved.reindex(self.data.index, fill_value=False), po_col_name] = (
                        extracted[resolved].values
                    )
                    self.data.loc[unresolved, po_col_name] = None

                    n_resolved   = int(resolved.sum())
                    n_unresolved = int(unresolved.sum())
                    if n_resolved:
                        print(f"  - Resolved {n_resolved} non-PO rows via CO Doc Line Item Txt.")
                    if n_unresolved:
                        print(f"  - {n_unresolved} rows have no PO in either column "
                              f"— will be logged to exceptions tab (MISSING_PO).")

            # ── ER PO cleanup (vectorised) ───────────────────────────────────────
            er_num_rx = r'(?i)\b(ER\d+)\b'
            er_mask   = self.data[type_col] == 'ER'

            if er_mask.any():
                po_col_name    = self.colmap['po']
                er_number_col  = 'ER Number' if 'ER Number' in self.data.columns else None
                er_sub         = self.data[er_mask].copy()

                # Start with the PO column itself as the base
                best = er_sub[po_col_name].astype(str).str.strip()

                # Build candidate columns in priority order
                candidate_cols: list[str] = []
                if er_number_col:
                    candidate_cols.append(er_number_col)
                candidate_cols += [c for c in (
                    'GL Line Description', 'GL Transaction Description',
                    'Description', 'CO Doc Line Item Txt', 'ER Description',
                ) if c in er_sub.columns]

                for col in candidate_cols:
                    extracted = (
                        er_sub[col].astype(str)
                        .str.extract(er_num_rx, expand=False)
                        .str.upper()
                    )
                    # Fill positions where we haven't found an ER number yet
                    needs_fill = ~best.str.match(r'(?i)^ER\d+$')
                    best = best.where(~needs_fill | extracted.isna(), extracted)

                self.data.loc[er_mask, po_col_name] = best.values
            print("Successfully loaded transactional data from valid sheets.")

        except Exception as e:
            print("Error loading transactional detail file:", e)

    def _categorize(self) -> 'pd.Series':
        """Categorise every row in self.data.

        Priority order:
        1. Document number / Vendor Invoice / AP Voucher Number prefix + amount sign
             "2" + positive → Accrual,  "2" + negative → Reversal
             "5" → Actual
             "9" → Reclass
        2. CO Doc Line Item Txt keywords (fallback when no classifier prefix matches):
             reversal → Reversal, accrual → Accrual,
             reclass  → Reclass,  invoice/vendor → Actual
        3. Amount sign for any row still unresolved:
             negative → Reversal, positive → Accrual, zero → Actual.
        """
        df     = self.data
        result = pd.Series(['Undefined'] * len(df), index=df.index, dtype=object)

        # ── Step 1: classifier prefix (Document number / Vendor Invoice / AP Voucher Number)
        cls_col = self.colmap.get('classifier', 'AP Voucher Number')
        if cls_col in df.columns:
            cls     = df[cls_col].astype(str).str.strip()
            amt_col = self.colmap.get('amount', 'GL BER Corp Amount')
            gl_col  = 'GL Transaction Amount'
            sign_col = gl_col if gl_col in df.columns else (amt_col if amt_col in df.columns else None)
            amt = pd.to_numeric(df[sign_col], errors='coerce').fillna(0) if sign_col else pd.Series(0.0, index=df.index)

            result[cls.str.startswith('5')] = 'Actual'
            result[cls.str.startswith('9')] = 'Reclass'

            two_mask = cls.str.startswith('2')
            result[two_mask & (amt >= 0)] = 'Accrual'
            result[two_mask & (amt <  0)] = 'Reversal'

        # ── Step 2: CO Doc Line Item Txt — fallback for rows with no prefix ──
        co_col = 'CO Doc Line Item Txt'
        if co_col in df.columns:
            desc  = df[co_col].astype(str).str.strip().str.lower()
            valid = ~desc.isin({'nan', 'none', ''})
            unset = result == 'Undefined'
            result[unset & valid & desc.str.contains('reversal', na=False)] = 'Reversal'
            unset = result == 'Undefined'
            result[unset & valid & desc.str.contains('accrual',  na=False)] = 'Accrual'
            unset = result == 'Undefined'
            result[unset & valid & desc.str.contains('reclass',  na=False)] = 'Reclass'
            unset = result == 'Undefined'
            result[unset & valid & (
                desc.str.contains('invoice', na=False) | desc.str.contains('vendor', na=False)
            )] = 'Actual'

        # ── Step 3: amount-sign fallback for anything still unresolved ───────
        unset = result == 'Undefined'
        if unset.any():
            amt_col = self.colmap.get('amount', 'GL BER Corp Amount')
            if amt_col in df.columns:
                amt = pd.to_numeric(df[amt_col], errors='coerce').fillna(0)
                result[unset & (amt <  0)] = 'Reversal'
                result[unset & (amt >  0)] = 'Accrual'
            result[result == 'Undefined'] = 'Actual'

        return result

    def get_transactional_data(self) -> dict:
        '''
        Method that gets transactional data from C-TIES file. 
        Extracts all rows, categorizes them, and returns a dict w/ actuals and accruals.
        Aggregates by PO / month.
        Returns:
            dict: {
                'PO12345': {
                    'cost_center': '1234',
                    'wbs': 'IT-CT123',
                    'Jan': {
                        'Actual': 900,
                        'Accrual': 950.0,
                        'Reversal': 0.0,
                    },
                    'Feb': {
                        'Actual': 800,
                        'Accrual': 1050.0,
                        'Reversal': -950,
                    }
                    ...
                }
                ...
            }
        '''
        # Load and preprocess (categorization happens during load)
        if self.data is None:
            self.load_transactional_detail_file()

        # Compute total Actuals amount per (PO, cost_center) so that the same PO
        # number appearing under different cost centers is not double-counted.
        ber_col  = self.colmap["amount"]
        po_col   = self.colmap["po"]
        cc_col   = self.colmap["cost_center"]
        type_col = self.colmap["type"]
        gross_by_po_cc: dict = {}  # keyed (po, cost_center)
        gross_by_po: dict = {}     # keyed po only — fallback for callers that don't have cc
        if ber_col in self.data.columns and type_col in self.data.columns:
            actuals_mask = self.data[type_col].isin(["Actual", "ER"])
            actuals_data = self.data.loc[actuals_mask].copy()
            actuals_data[ber_col] = pd.to_numeric(actuals_data[ber_col], errors='coerce').fillna(0)
            gross_by_po = actuals_data.groupby(po_col)[ber_col].sum().to_dict()
            if cc_col in actuals_data.columns:
                gross_by_po_cc = (
                    actuals_data.groupby([po_col, cc_col])[ber_col].sum()
                    .reset_index()
                    .set_index([po_col, cc_col])[ber_col]
                    .to_dict()
                )

        # Filter to only the columns needed
        country_col = self.colmap.get("country")
        if country_col and country_col not in self.data.columns:
            country_col = None

        posting_date_col = self.colmap.get("gl_posting_date")
        if posting_date_col and posting_date_col not in self.data.columns:
            posting_date_col = None

        cols = [
            self.colmap["po"],
            self.colmap["month"],
            self.colmap["amount"],
            self.colmap["cost_center"],
            self.colmap["wbs"],
            self.colmap["type"],
        ]
        if country_col:
            cols.append(country_col)
        if posting_date_col:
            cols.append(posting_date_col)
        missing = [c for c in cols if c not in self.data.columns]
        if missing:
            raise ValueError(
                f"Missing required columns in the loaded transactional dataset: {missing}. "
                f"These columns are required for the transactional aggregation process. "
                f"Please verify that your config colmap and raw sheet columns align. "
                f"Available columns in the dataset after normalization: {list(self.data.columns)}"
            )
        data_copy = self.data[cols].copy()
        # Ensure amount column is numeric
        data_copy[self.colmap["amount"]] = pd.to_numeric(
            data_copy[self.colmap["amount"]],
            errors='coerce'
        )
        # Filter valid transaction types, and also include Reclass rows so they
        # contribute to the PO's Actual total.
        valid_with_reclass = list(self.valid_types) + ["Reclass"]
        data_copy = data_copy[data_copy[self.colmap["type"]].isin(valid_with_reclass)]

        # Derive the current fiscal year from the AP values in the data so that
        # two December rows (AP12 2025 = prior-year Dec, AP12 2026 = current Dec)
        # can be told apart.  The maximum year seen across all AP values is treated
        # as the current year; any December with a smaller year is "Dec (PY)".
        month_col = self.colmap["month"]
        parsed_years = data_copy[month_col].apply(lambda v: self._parse_ap(v)[1])
        known_years = parsed_years.dropna()
        current_year = int(known_years.max()) if not known_years.empty else None

        # Group by PO / Month / Type / Cost Center / WBS
        group_cols = [
            self.colmap["po"],
            self.colmap["month"],
            self.colmap["type"],
            self.colmap["cost_center"],
            self.colmap["wbs"],
        ]
        grouped = (
            data_copy.groupby(group_cols)[self.colmap["amount"]]
            .sum()
            .reset_index()
        )

        # Build results.
        # Month mapping is direct: AP01 = Jan, AP02 = Feb, … AP12 = Dec.
        # The only special case is December — when two Decembers appear in the
        # data (AP12 of the prior year and AP12 of the current year), the one
        # whose year is less than current_year lands in "Dec (PY)".
        result = {}
        for _, row in grouped.iterrows():
            po          = row[self.colmap["po"]]
            raw_month   = row[self.colmap["month"]]
            type_name   = row[self.colmap["type"]]
            value       = row[self.colmap["amount"]]
            cost_center = str(row[self.colmap["cost_center"]]).strip()
            wbs         = str(row[self.colmap["wbs"]]).strip()

            month_num, ap_year = self._parse_ap(raw_month)
            if month_num is None:
                continue

            # Actuals (Actual, ER, Reclass) and Reversals shift back one month:
            # AP posted in period N reflects activity from period N-1.
            # AP01 → "Dec (25)" (the prior-year December column), AP02 → Jan, … AP12 → Nov.
            #
            # Accruals map directly: AP01 = Jan, AP02 = Feb, … AP12 = Dec.
            # Prior-year data (AP12 of a year < current_year) is skipped.
            if type_name in ("Actual", "ER", "Reclass", "Reversal"):
                shifted_num = month_num - 1
                if shifted_num == 0:
                    # AP01 of the current year → prior-year December column.
                    # The year suffix matches the prior year (current_year - 1), shortened to 2 digits.
                    # e.g. current_year=2026 → "Dec (25)"; current_year=2027 → "Dec (26)"
                    py = (current_year - 1) if current_year else None
                    write_month = f"Dec ({str(py)[2:]})" if py else "Dec"
                else:
                    write_month = self.month_map.get(shifted_num)
            elif month_num == 12 and current_year is not None and ap_year is not None and ap_year < current_year:
                continue  # prior-year December Accrual — skip, same as any prior-year data
            else:
                write_month = self.month_map.get(month_num)

            if not write_month:
                continue

            # Initialize (PO, cost_center) bucket
            key = (po, cost_center)
            if key not in result:
                result[key] = {
                    "cost_center": cost_center,
                    "wbs": wbs,
                    "gross_ber_total": gross_by_po_cc.get(key) or gross_by_po.get(po, 0.0),
                }

            # Initialize month bucket
            if write_month not in result[key]:
                result[key][write_month] = {
                    "Actual": 0,
                    "Accrual": 0,
                    "Reversal": 0,
                }

            # Assign value by type
            if type_name in ["Actual", "ER", "Reclass"]:
                result[key][write_month]["Actual"] = result[key][write_month].get("Actual", 0) + value
            elif type_name in ["Accrual", "Reversal"]:
                result[key][write_month][type_name] = value

        # Sort months chronologically: prior-year December suffix first, then Jan–Dec.
        # The prior-year December key is dynamic e.g. "Dec (25)" — collect any such
        # keys present in any PO bucket and sort them before January.
        _py_dec_keys = sorted({
            m for vals in result.values()
            for m in vals
            if isinstance(m, str) and m.startswith("Dec (")
        })
        month_order = _py_dec_keys + [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ]
        for key in result:
            result[key] = {
                "cost_center": result[key]["cost_center"],
                "wbs": result[key]["wbs"],
                "gross_ber_total": result[key].get("gross_ber_total", 0.0),
                **{
                    month: result[key][month]
                    for month in month_order
                    if month in result[key]
                }
            }

        # DEBUG — print first 3 PO keys so month placement is visible
        print("DEBUG transactional_data sample:")
        for i, (key, val) in enumerate(result.items()):
            if i >= 3:
                break
            months = {m: v for m, v in val.items() if m not in ('cost_center', 'wbs', 'gross_ber_total')}
            print(f"  {key}: {months}")

        return result
    


    def get_intl_po_set(self) -> set:
        """Return the set of PO numbers whose country marks them as international.

        Uses the same _INTL_COUNTRIES check as get_transactional_data() so the
        two are always consistent.  Result is derived from the raw loaded data so
        it includes every row (Actuals, Accruals, Reclasses, etc.).
        """
        if self.data is None:
            self.load_transactional_detail_file()

        country_col = self.colmap.get("country")
        if not country_col or country_col not in self.data.columns:
            return set()

        _INTL_COUNTRIES = {'india', 'in'}
        po_col = self.colmap["po"]
        cc_data = self.data[[po_col, country_col]].copy()
        cc_data[country_col] = cc_data[country_col].astype(str).str.strip().str.lower()
        intl = cc_data[cc_data[country_col].isin(_INTL_COUNTRIES)]
        return set(intl[po_col].astype(str).str.strip().unique())


    def get_reclass_data(self) -> dict:
        """
        Aggregates Reclass rows by cost center and month.

        Reclass rows have no real PO or WBS — they carry a 9xx AP voucher number
        in the PO column and '#' in the WBS column.  They should be written to the
        template as a single dedicated row per cost center (labeled 'RECLASS-{cc}')
        rather than being merged into a PO row.

        Returns:
            dict: {
                '1234': {
                    'Jan': 500.0,
                    'Feb': -200.0,
                    ...
                },
                ...
            }
        """
        if self.data is None:
            self.load_transactional_detail_file()

        reclass_rows = self.data[self.data[self.colmap["type"]] == "Reclass"].copy()
        if reclass_rows.empty:
            return {}

        reclass_rows[self.colmap["amount"]] = pd.to_numeric(
            reclass_rows[self.colmap["amount"]], errors='coerce'
        ).fillna(0.0)

        _parsed_years = reclass_rows[self.colmap["month"]].apply(lambda v: self._parse_ap(v)[1])
        _current_year = int(_parsed_years.dropna().max()) if not _parsed_years.dropna().empty else None

        result = {}
        for _, row in reclass_rows.iterrows():
            cc = str(row[self.colmap["cost_center"]]).strip()
            if not cc or cc.lower() in ('#', 'nan', 'none', ''):
                continue

            raw_month = row[self.colmap["month"]]
            month_num, _ = self._parse_ap(raw_month)
            if month_num is None:
                continue

            # Shift back one month: AP01 → "Dec (YY)", AP02 → Jan, …
            # e.g. current_year=2026 → "Dec (25)"; current_year=2027 → "Dec (26)"
            shifted_num = month_num - 1
            if shifted_num == 0:
                py = (_current_year - 1) if _current_year else None
                month_label = f"Dec ({str(py)[2:]})" if py else "Dec"
            else:
                month_label = self.month_map.get(shifted_num)
            if not month_label:
                continue

            amount = float(row[self.colmap["amount"]])
            if cc not in result:
                result[cc] = {}
            result[cc][month_label] = result[cc].get(month_label, 0.0) + amount

        print(f"Reclass data aggregated for {len(result)} cost center(s).")
        return result


    def get_reclass_notes(self) -> dict:
        """
        Returns per-PO reclass note data for PO-linked reclasses (Reclass_PO rows).
        Used by template_writer to annotate the Actual cell with a comment and highlight.

        Returns:
            dict: {
                'PO12345': {
                    'Jan': [(amount, description), ...],
                    ...
                },
                ...
            }
        """
        if self.data is None:
            self.load_transactional_detail_file()

        reclass_po_rows = self.data[self.data[self.colmap["type"]] == "Reclass"].copy()
        if reclass_po_rows.empty:
            return {}

        reclass_po_rows[self.colmap["amount"]] = pd.to_numeric(
            reclass_po_rows[self.colmap["amount"]], errors='coerce'
        ).fillna(0.0)

        _rn_parsed_years = reclass_po_rows[self.colmap["month"]].apply(lambda v: self._parse_ap(v)[1])
        _rn_current_year = int(_rn_parsed_years.dropna().max()) if not _rn_parsed_years.dropna().empty else None

        result = {}
        for _, row in reclass_po_rows.iterrows():
            po = str(row[self.colmap["po"]]).strip()
            if not po or po.lower() in ('#', 'nan', 'none', ''):
                continue

            raw_month = row[self.colmap["month"]]
            month_num, _ = self._parse_ap(raw_month)
            if month_num is None:
                continue

            # Shift back one month: AP01 → "Dec (YY)", AP02 → Jan, …
            # e.g. current_year=2026 → "Dec (25)"; current_year=2027 → "Dec (26)"
            shifted_num = month_num - 1
            if shifted_num == 0:
                py = (_rn_current_year - 1) if _rn_current_year else None
                month_label = f"Dec ({str(py)[2:]})" if py else "Dec"
            else:
                month_label = self.month_map.get(shifted_num)

            po_key = po
            if po_key not in result and po_key.replace('.', '', 1).replace('-', '', 1).isdigit():
                try:
                    po_key = str(int(float(po_key)))
                except (ValueError, OverflowError):
                    po_key = po
            if not month_label:
                continue

            amount = float(row[self.colmap["amount"]])

            # Pick the best description for the note
            description = ""
            for desc_col in ("Description", "GL Line Description", "GL Transaction Description"):
                val = str(row.get(desc_col, "")).strip()
                if val and val.lower() not in ('nan', 'none', ''):
                    description = val
                    break

            if po_key not in result:
                result[po_key] = {}
            if month_label not in result[po_key]:
                result[po_key][month_label] = []
            result[po_key][month_label].append((amount, description))

        print(f"Reclass notes collected for {len(result)} PO(s).")
        return result



    
    def get_hierarchy_map(self) -> dict:
        """
        Returns a mapping of every row in the transactional file,
        keyed by row index. Ensures every row is accounted for.
        Used for hierarchy building and exception tracking.
        
        Special handling: Reads GL Line Description, GL Transaction Description,
        and Description columns for ER number extraction.
        
        Returns:
            dict: {
                45: { 'po': 'PO1234', 'cost_center': '1234', 'wbs': 'IT-CT123',
                      'gl_line_desc': '...', 'gl_trans_desc': '...', 'description': '...' },
                46: { 'po': None, 'cost_center': '2345', 'wbs': None,
                      'gl_line_desc': '...', 'gl_trans_desc': None, 'description': None },
                ...
            }
        """
        if self.data is None:
            self.load_transactional_detail_file()
        # Validate columns exist before iterating
        required = [self.colmap["po"], self.colmap["wbs"], self.colmap["cost_center"]]
        missing_cols = [c for c in required if c not in self.data.columns]
        if missing_cols:
            raise ValueError(
                f"Missing required columns for building the hierarchy map: {missing_cols}. "
                f"These fields (PO Number, WBS Element, and Cost Center) must be mapped to valid columns in your dataset. "
                f"Please inspect config mapping or sheet header row. "
                f"Available columns in normalized dataset: {list(self.data.columns)}"
            )

        # Legal entity column is optional — gracefully absent when not in colmap or file
        le_col = self.colmap.get("legal_entity")
        if le_col and le_col not in self.data.columns:
            le_col = None

        # Country column is optional — gracefully absent when not in colmap or file
        country_col = self.colmap.get("country")
        if country_col and country_col not in self.data.columns:
            country_col = None

        # Vendor name column is optional — gracefully absent when not in colmap or file
        vendor_col = self.colmap.get("vendor_name")
        if vendor_col and vendor_col not in self.data.columns:
            vendor_col = None

        # Pre-compute the best (most frequent real) vendor name per PO.
        # Uses case-insensitive filtering so "Not assigned", "NOT ASSIGNED", etc. are all excluded.
        _PLACEHOLDER_LOWER = {'', 'nan', 'none', '#', 'not assigned'}
        vendor_by_po: dict = {}
        if vendor_col:
            _po_col = self.colmap["po"]
            _vc = self.data[[_po_col, vendor_col]].copy()
            _vc[vendor_col] = _vc[vendor_col].astype(str).str.strip()
            # Keep only rows whose vendor is a real value (case-insensitive)
            _vc = _vc[~_vc[vendor_col].str.lower().isin(_PLACEHOLDER_LOWER)]
            if not _vc.empty:
                # Cast PO key to string to match the normalized po used in lookup
                _vc = _vc.copy()
                _vc[_po_col] = _vc[_po_col].astype(str).str.strip()
                vendor_by_po = (
                    _vc.groupby(_po_col)[vendor_col]
                    .agg(lambda s: s.mode().iloc[0] if not s.mode().empty else None)
                    .to_dict()
                )

        # G/L Account column is optional — gracefully absent when not in colmap or file
        gl_account_col = self.colmap.get("gl_account")
        if gl_account_col and gl_account_col not in self.data.columns:
            gl_account_col = None

        # Requisition title column is optional — gracefully absent when not in colmap or file.
        # Try the configured name first, then fall back through common alternate names.
        req_title_col = self.colmap.get("req_title")
        if not req_title_col or req_title_col not in self.data.columns:
            for _fallback in (
                "GL Line Description",
                "GL Transaction Description",
                "GL Description",
                "Description",
                "PO Line Item Desc",
            ):
                if _fallback in self.data.columns:
                    req_title_col = _fallback
                    break
            else:
                req_title_col = None

        # Project Name column is optional — gracefully absent when not in colmap or file
        project_name_col = self.colmap.get("project_name")
        if project_name_col and project_name_col not in self.data.columns:
            project_name_col = None

        # Pre-compute the best (most frequent real) project name per PO.
        # Using mode aggregation ensures that even when a PO's first occurrence in the
        # hierarchy map has no project name (e.g. a Reclass row), the correct name from
        # other rows for that PO is still picked up — exactly like vendor_by_po.
        project_name_by_po: dict = {}
        if project_name_col:
            _po_col = self.colmap["po"]
            _pnc = self.data[[_po_col, project_name_col]].copy()
            _pnc[project_name_col] = _pnc[project_name_col].astype(str).str.strip()
            _pnc = _pnc[~_pnc[project_name_col].str.lower().isin(_PLACEHOLDER_LOWER)]
            if not _pnc.empty:
                _pnc = _pnc.copy()
                _pnc[_po_col] = _pnc[_po_col].astype(str).str.strip()
                project_name_by_po = (
                    _pnc.groupby(_po_col)[project_name_col]
                    .agg(lambda s: s.mode().iloc[0] if not s.mode().empty else None)
                    .to_dict()
                )

        # Pre-compute the most frequent real req_title per PO (mode).
        # Uses case-insensitive filtering so "Not assigned" etc. are excluded.
        # PO key is cast to string+strip to match the normalized po used in the lookup.
        req_title_by_po: dict = {}
        if req_title_col:
            po_col = self.colmap["po"]
            cleaned = self.data[[po_col, req_title_col]].copy()
            cleaned[po_col] = cleaned[po_col].astype(str).str.strip()
            cleaned[req_title_col] = cleaned[req_title_col].astype(str).str.strip()
            cleaned = cleaned[~cleaned[req_title_col].str.lower().isin(_PLACEHOLDER_LOWER)]
            if not cleaned.empty:
                req_title_by_po = (
                    cleaned.groupby(po_col)[req_title_col]
                    .agg(lambda s: s.mode().iloc[0] if not s.mode().empty else None)
                    .to_dict()
                )

        # ── Vectorised build — no iterrows ───────────────────────────────────
        _PLACEHOLDER_LOWER = {'', 'none', 'nan', '#', 'not assigned'}

        def _clean_series(s: pd.Series) -> pd.Series:
            """Strip, cast to str, replace placeholders with None (as object)."""
            cleaned = s.astype(str).str.strip()
            cleaned[cleaned.str.lower().isin(_PLACEHOLDER_LOWER)] = None
            cleaned[s.isna()] = None
            return cleaned

        po_s  = _clean_series(self.data[self.colmap["po"]])
        wbs_s = _clean_series(self.data[self.colmap["wbs"]])
        cc_s  = _clean_series(self.data[self.colmap["cost_center"]])

        # Legal entity — strip trailing .0 from numeric values
        if le_col:
            le_raw = self.data[le_col]
            le_num = pd.to_numeric(le_raw, errors='coerce')
            le_s   = le_num.dropna().astype('int64').astype(str)
            le_s   = le_s.reindex(self.data.index)
            # Fill non-numeric positions with the original string value
            le_fallback = le_raw.astype(str).str.strip()
            le_s = le_s.fillna(le_fallback)
            le_s[le_s.str.lower().isin(_PLACEHOLDER_LOWER) | le_raw.isna()] = None
        else:
            le_s = pd.Series([None] * len(self.data), index=self.data.index, dtype=object)

        # Country
        if country_col:
            country_s = _clean_series(self.data[country_col])
        else:
            country_s = pd.Series([None] * len(self.data), index=self.data.index, dtype=object)

        # Project Name — map pre-computed best value via PO key (same as vendor_name)
        # so the correct name is returned even when a PO's first row lacks the field.
        if project_name_by_po:
            project_name_s = po_s.map(project_name_by_po)
        else:
            project_name_s = pd.Series([None] * len(self.data), index=self.data.index, dtype=object)

        # Vendor — map pre-computed best value via PO key
        if vendor_by_po:
            vendor_s = po_s.map(vendor_by_po)
        else:
            vendor_s = pd.Series([None] * len(self.data), index=self.data.index, dtype=object)

        # G/L account — strip trailing .0
        if gl_account_col:
            gl_raw = self.data[gl_account_col]
            gl_num = pd.to_numeric(gl_raw, errors='coerce')
            gl_s   = gl_num.dropna().astype('int64').astype(str)
            gl_s   = gl_s.reindex(self.data.index)
            gl_fallback = gl_raw.astype(str).str.strip()
            gl_s = gl_s.fillna(gl_fallback)
            gl_s[gl_s.str.lower().isin(_PLACEHOLDER_LOWER) | gl_raw.isna()] = None
        else:
            gl_s = pd.Series([None] * len(self.data), index=self.data.index, dtype=object)

        # Requisition title — map pre-computed mode via PO key
        if req_title_by_po:
            req_s = po_s.map(req_title_by_po)
        else:
            req_s = pd.Series([None] * len(self.data), index=self.data.index, dtype=object)

        # Description columns
        desc_col_names = {
            'gl_line_desc':  'GL Line Description',
            'gl_trans_desc': 'GL Transaction Description',
            'description':   'Description',
        }
        desc_series: dict[str, pd.Series] = {}
        for key, col in desc_col_names.items():
            if col in self.data.columns:
                s = self.data[col].astype(str).str.strip()
                s[s.str.lower().isin({'', 'nan', 'none'})] = None
                s[self.data[col].isna()] = None
                desc_series[key] = s
            else:
                desc_series[key] = pd.Series([None] * len(self.data),
                                             index=self.data.index, dtype=object)

        # Assemble into list-of-dicts then convert to indexed dict
        # zip over numpy arrays is ~10× faster than iterrows
        # Include the 1-based Excel row number so the Exceptions tab shows the
        # correct row that can be found directly in the source file.
        excel_row_s = (
            self._excel_row
            if self._excel_row is not None
            else (self.data.index + 2)  # fallback: assume header on row 1
        )
        records = list(zip(
            self.data.index,
            excel_row_s,
            po_s,  wbs_s, cc_s, le_s, country_s,
            vendor_s, gl_s, req_s,
            desc_series['gl_line_desc'],
            desc_series['gl_trans_desc'],
            desc_series['description'],
            project_name_s,
        ))
        result = {
            idx: {
                'po':           po   if po   != 'None' and po   is not None else None,
                'wbs':          wbs  if wbs  != 'None' and wbs  is not None else None,
                'cost_center':  cc   if cc   != 'None' and cc   is not None else None,
                'legal_entity': le   if le   != 'None' and le   is not None else None,
                'country':      ctr  if ctr  != 'None' and ctr  is not None else None,
                'vendor_name':  vnd  if isinstance(vnd, str) and vnd else None,
                'gl_account':   gl   if gl   != 'None' and gl   is not None else None,
                'req_title':    req  if isinstance(req, str) and req else None,
                'gl_line_desc': gld  if gld  != 'None' and gld  is not None else None,
                'gl_trans_desc':gtd  if gtd  != 'None' and gtd  is not None else None,
                'description':  dsc  if dsc  != 'None' and dsc  is not None else None,
                'project_name': prj  if prj  != 'None' and prj  is not None else None,
                'excel_row':    int(xrow) if xrow is not None else None,
            }
            for idx, xrow, po, wbs, cc, le, ctr, vnd, gl, req, gld, gtd, dsc, prj in records
        }

        print(f"Hierarchy map built: {len(result)} rows processed.")
        print(f"  - Missing PO:          {sum(1 for v in result.values() if v['po'] is None)}")
        print(f"  - Missing WBS:         {sum(1 for v in result.values() if v['wbs'] is None)}")
        print(f"  - Missing Cost Center: {sum(1 for v in result.values() if v['cost_center'] is None)}")
        return result