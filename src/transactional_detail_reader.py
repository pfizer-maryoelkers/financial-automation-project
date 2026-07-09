
import pandas as pd
import re

class TransactionalDetailReader:
    """Class for reading transactional detail file and extracts accruals, actuals, and reversals.

    Includes the following methods:
        - load_transactional_detail_file(): Loads TIES file into singular dataframe
        - _categorize_row(): categorizes row type as actual, accrual, etc.
        - get_transactional_data(): reads dataframe and filters data, returns dict of data we need
        
    Codes for AP Voucher Number (used in categorize row):
        210 - Accrual/Reversal
        510 - Invoice
        900 - Reclass

    __init__ defines required cols, required types, and a column map for easy configuration of newly formatted CTIES files.
    
    """

    def __init__(self, file_path, required_cols, valid_types, colmap):
        """Initialize with the transactional detail file path.
        
        See config_base.yaml for default parameters
        """
        self.file_path = file_path
        self.data = None

        self.required_cols = set(required_cols) # Required columns for loading a valid sheet
        
        self.valid_types = valid_types # Valid types for reading (currently support Actuals, Accruals, Reversals) - open to extension

        self.colmap = colmap # Column map for standardized writing. Can change values for dynamic column names

        self.month_map = {
            1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
            7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
        } # Month map for reading



    def load_transactional_detail_file(self):
        """Load all valid sheets from the transactional Excel file.
        Valid sheet defined by having the following columns: 'PO Number', 'Month', 'GL Transaction Amount'
        """
        try:
            xls = pd.ExcelFile(self.file_path)

            valid_sheets = []
            for sheet in xls.sheet_names:
                # Read just a preview to check column names
                preview = pd.read_excel(self.file_path, sheet_name=sheet, header=1, nrows=5)

                if self.required_cols.issubset(preview.columns):
                    valid_sheets.append(sheet)

            if not valid_sheets:
                raise ValueError("No valid sheets found containing required transactional columns.")

            print(f"Loading valid sheets: {valid_sheets}")

            dfs = [
                pd.read_excel(self.file_path, sheet_name=sheet, header=1)
                for sheet in valid_sheets
            ]

            self.data = pd.concat(dfs, ignore_index=True)
            self.data["Type"] = self.data.apply(self._categorize_row, axis=1)
            self.data[self.colmap['po']] = self.data[self.colmap['po']].astype('str') # Change POs to strings

            er_num_pattern = re.compile(r'\bER\d+\b', re.IGNORECASE)
            er_mask = self.data["Type"] == "ER"
            if er_mask.any():
                def _clean_er_po(row):
                    po_val = str(row[self.colmap['po']]).strip()
                    match = er_num_pattern.search(po_val)
                    if match:
                        return match.group(0).upper()
                    for col in ("GL Line Description", "GL Transaction Description", "Description", "CO Doc Line Item Txt"):
                        if col in row.index:
                            txt = str(row[col]).strip()
                            if txt and txt.lower() not in ('nan', 'none', ''):
                                match = er_num_pattern.search(txt)
                                if match:
                                    return match.group(0).upper()
                    return po_val

                self.data.loc[er_mask, self.colmap['po']] = (
                    self.data[er_mask].apply(_clean_er_po, axis=1)
                )
            print("Successfully loaded transactional data from valid sheets.")

        except Exception as e:
            print("Error loading transactional detail file:", e)

    # Internal helper function that categorizes a row in CTIES file as Actual, Accrual, Reversal, etc.
    # Use by calling df["Type"] = df.apply(self._categorize_row, axis=1)
    def _categorize_row(self, row):
        '''
        Returns value for row 'Type' as a string.

        Priority order:
        1. 9xx vouchers:
             - Any description containing ER<digits> = ER
             - otherwise Reclass
        2. "CO Doc Line Item Txt" description column — checked for keywords
           (accrual, reversal, reclass, invoice) as the most reliable source.
        3. AP Voucher Number prefix as fallback:
             "5xx" = Actual (vendor invoice)
             "2xx" = Accrual (positive GL Transaction Amount) or Reversal (negative)
             "9xx" = Reclass
        '''
        classifier = str(row[self.colmap["classifier"]])

        if classifier.startswith("9"):
            for desc_col in ("GL Line Description", "GL Transaction Description", "Description", "CO Doc Line Item Txt"):
                desc = str(row.get(desc_col, "")).strip()
                if desc and desc.lower() not in ('nan', 'none', '') and re.search(r'\bER\d+\b', desc, re.IGNORECASE):
                    return "ER"
            return "Reclass"

        # --- Step 1: check CO Doc Line Item Txt for explicit description ---
        co_doc_col = "CO Doc Line Item Txt"
        if co_doc_col in row.index:
            desc = str(row[co_doc_col]).strip().lower()
            if desc and desc not in ('nan', 'none', ''):
                if 'reversal' in desc:
                    return "Reversal"
                if 'accrual' in desc:
                    return "Accrual"
                if 'reclass' in desc:
                    return "Reclass"
                if 'invoice' in desc or 'vendor' in desc:
                    return "Actual"

        # --- Step 2: fall back to AP Voucher Number prefix ---
        # Use GL Transaction Amount for sign — always populated.
        # Fall back to the configured amount column if the column is absent.
        gl_trans_col = "GL Transaction Amount"
        if gl_trans_col in row.index:
            try:
                sign_amount = float(row[gl_trans_col])
            except (TypeError, ValueError):
                sign_amount = 0.0
        else:
            try:
                sign_amount = float(row[self.colmap["amount"]])
            except (TypeError, ValueError):
                sign_amount = 0.0

        if classifier.startswith("5"):
            return "Actual"
        elif classifier.startswith("2"):
            if sign_amount >= 0:
                return "Accrual"
            else:
                return "Reversal"
        else:
            return "Undefined"

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
        # Filter to only the columns needed
        cols = [
            self.colmap["po"],
            self.colmap["month"],
            self.colmap["amount"],
            self.colmap["cost_center"],
            self.colmap["wbs"],
            self.colmap["type"]
        ]
        missing = [c for c in cols if c not in self.data.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}")
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
        # Group by PO / Month / Type / Cost Center / WBS
        grouped = (
            data_copy.groupby([
                self.colmap["po"],
                self.colmap["month"],
                self.colmap["type"],
                self.colmap["cost_center"],
                self.colmap["wbs"]
            ])[self.colmap["amount"]]
            .sum()
            .reset_index()
        )
        # Build results
        result = {}
        for _, row in grouped.iterrows():
            po = row[self.colmap["po"]]
            raw_month = row[self.colmap["month"]]
            type_name = row[self.colmap["type"]]
            value = row[self.colmap["amount"]]
            cost_center = str(row[self.colmap["cost_center"]]).strip()
            wbs = str(row[self.colmap["wbs"]]).strip()

            # Normalise month to 1-12.
            # Supports plain integers (1-12) and YYYYMM format (e.g. 202601 → 1).
            try:
                raw_month_int = int(raw_month)
                if raw_month_int > 12:
                    # YYYYMM format — extract last two digits
                    month_num = raw_month_int % 100
                else:
                    month_num = raw_month_int
            except (TypeError, ValueError):
                continue  # unparseable month — skip row

            # Actual values belong to prior month
            if month_num == 1:
                actual_month = "Dec"
            else:
                actual_month = self.month_map.get(month_num - 1)
            # Accruals/reversals belong to current month
            accrual_month = self.month_map.get(month_num)
            # Initialize PO
            if po not in result:
                result[po] = {
                    "cost_center": cost_center,
                    "wbs": wbs
                }
            # Initialize month bucket
            if type_name in ["Actual", "ER", "Reclass"] and actual_month not in result[po]:
                result[po][actual_month] = {
                    "Actual": 0,
                    "Accrual": 0,
                    "Reversal": 0,
                }
            elif type_name in ["Accrual", "Reversal"] and accrual_month not in result[po]:
                result[po][accrual_month] = {
                    "Actual": 0,
                    "Accrual": 0,
                    "Reversal": 0,
                }
            # Assign value
            if type_name in ["Actual", "ER", "Reclass"]:
                result[po][actual_month]["Actual"] = result[po][actual_month].get("Actual", 0) + value
            elif type_name in ["Accrual", "Reversal"] and accrual_month:
                result[po][accrual_month][type_name] = value
        # Sort months for readability, preserve cost_center and wbs
        month_order = [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ]
        for po in result:
            result[po] = {
                "cost_center": result[po]["cost_center"],
                "wbs": result[po]["wbs"],
                **{
                    month: result[po][month]
                    for month in month_order
                    if month in result[po]
                }
            }
        return result
    

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

        result = {}
        for _, row in reclass_rows.iterrows():
            cc = str(row[self.colmap["cost_center"]]).strip()
            if not cc or cc.lower() in ('#', 'nan', 'none', ''):
                continue

            raw_month = row[self.colmap["month"]]
            try:
                raw_month_int = int(raw_month)
                month_num = raw_month_int % 100 if raw_month_int > 12 else raw_month_int
            except (TypeError, ValueError):
                continue

            # Reclass amounts are posted in the current period — use actual_month
            # (current month - 1) to stay consistent with how Actuals are placed.
            if month_num == 1:
                month_label = "Dec"
            else:
                month_label = self.month_map.get(month_num - 1)
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

        result = {}
        for _, row in reclass_po_rows.iterrows():
            po = str(row[self.colmap["po"]]).strip()
            if not po or po.lower() in ('#', 'nan', 'none', ''):
                continue

            raw_month = row[self.colmap["month"]]
            try:
                raw_month_int = int(raw_month)
                month_num = raw_month_int % 100 if raw_month_int > 12 else raw_month_int
            except (TypeError, ValueError):
                continue

            if month_num == 1:
                month_label = "Dec"
            else:
                month_label = self.month_map.get(month_num - 1)
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

            if po not in result:
                result[po] = {}
            if month_label not in result[po]:
                result[po][month_label] = []
            result[po][month_label].append((amount, description))

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
            raise ValueError(f"Missing required columns for hierarchy map: {missing_cols}")
        
        # Description columns to search for ER numbers (check all that exist in the file)
        desc_col_names = {
            'gl_line_desc': 'GL Line Description',
            'gl_trans_desc': 'GL Transaction Description',
            'description': 'Description',
        }
        present_desc_cols = {key: col for key, col in desc_col_names.items() if col in self.data.columns}
        
        # Placeholder values to treat as missing
        MISSING_VALUES = {'', 'none', 'nan', '#'}
        result = {}
        for idx, row in self.data.iterrows():
            po = row[self.colmap["po"]]
            wbs = row[self.colmap["wbs"]]
            cost_center = row[self.colmap["cost_center"]]
            
            # Normalize PO
            po = str(po).strip() if pd.notna(po) else None
            if po and po.lower() in MISSING_VALUES:
                po = None
            
            # Normalize WBS
            wbs = str(wbs).strip() if pd.notna(wbs) else None
            if wbs and wbs.lower() in MISSING_VALUES:
                wbs = None
            
            # Normalize cost center
            cost_center = str(cost_center).strip() if pd.notna(cost_center) else None
            if cost_center and cost_center.lower() in MISSING_VALUES:
                cost_center = None
            
            # Read all description columns for ER extraction
            desc_values = {}
            for key, col in present_desc_cols.items():
                val = row.get(col)
                desc_values[key] = str(val).strip() if pd.notna(val) and str(val).strip() else None
            
            result[idx] = {
                'po': po,
                'cost_center': cost_center,
                'wbs': wbs,
                'gl_line_desc': desc_values.get('gl_line_desc'),
                'gl_trans_desc': desc_values.get('gl_trans_desc'),
                'description': desc_values.get('description'),
            }
        
        # Ensure no rows were dropped
        assert len(result) == len(self.data), (
            f"Row count mismatch: expected {len(self.data)} rows, "
            f"got {len(result)}. Some rows may have been lost."
        )
        print(f"Hierarchy map built: {len(result)} rows processed.")
        print(f"  - Missing PO:          {sum(1 for v in result.values() if v['po'] is None)}")
        print(f"  - Missing WBS:         {sum(1 for v in result.values() if v['wbs'] is None)}")
        print(f"  - Missing Cost Center: {sum(1 for v in result.values() if v['cost_center'] is None)}")
        return result