
import pandas as pd

class TransactionalDetailReader:
    """Class for reading transactional detail file and extracts accruals, actuals, and reversals.

    Includes the following methods:
        - load_transactional_detail_file(): Loads CTIES file into singular dataframe
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
            print("Successfully loaded transactional data from valid sheets.")

        except Exception as e:
            print("Error loading transactional detail file:", e)

    # Internal helper function that categorizes a row in CTIES file as Actual, Accrual, Reversal, etc.
    # Use by calling df["Type"] = df.apply(self._categorize_row, axis=1)
    def _categorize_row(self, row):
        '''
        Returns value for row 'Type' as a string
        '''
        classifier = str(row[self.colmap["classifier"]])
        amount = row[self.colmap["amount"]]

        if classifier.startswith("5"):
            return "Actual"
        elif classifier.startswith("2"):
            if amount > 0:
                return "Accrual"
            elif amount < 0:
                return "Reversal"
        elif classifier.startswith("9"):
            return "Reclass"
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
        # Filter valid transaction types
        data_copy = data_copy[data_copy[self.colmap["type"]].isin(self.valid_types)]
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
            month_num = row[self.colmap["month"]]
            type_name = row[self.colmap["type"]]
            value = row[self.colmap["amount"]]
            cost_center = str(row[self.colmap["cost_center"]]).strip()
            wbs = str(row[self.colmap["wbs"]]).strip()
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
            if type_name == "Actual" and actual_month not in result[po]:
                result[po][actual_month] = {
                    "Actual": 0,
                    "Accrual": 0,
                    "Reversal": 0
                }
            elif type_name in ["Accrual", "Reversal"] and accrual_month not in result[po]:
                result[po][accrual_month] = {
                    "Actual": 0,
                    "Accrual": 0,
                    "Reversal": 0
                }
            # Assign value
            if type_name == "Actual":
                result[po][actual_month][type_name] = value
            else:
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
    
    
    def get_hierarchy_map(self) -> dict:
        """
        Returns a mapping of every row in the transactional file,
        keyed by row index. Ensures every row is accounted for.
        Used for hierarchy building and exception tracking.
        Returns:
            dict: {
                45: { 'po': 'PO1234', 'cost_center': '1234', 'wbs': 'IT-CT123' },
                46: { 'po': None,     'cost_center': '2345', 'wbs': None },
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
            result[idx] = {
                'po': po,
                'cost_center': cost_center,
                'wbs': wbs
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