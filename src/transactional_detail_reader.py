
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

    def __init__(self, file_path: str):
        """Initialize with the transactional detail file path."""
        self.file_path = file_path
        self.data = None

        self.required_cols = {
            'PO Number', 
            'Month', 
            'GL Transaction Amount'
         } # Required columns for loading a valid sheet
        
        self.valid_types =  [
            'Actual', 
            'Accrual', 
            'Reversal'
        ] # Valid types for reading (currently support Actuals, Accruals, Reversals) - open to extension

        self.colmap = {
            "po": "PO Number",
            "month": "Month",
            "amount": "GL Transaction Amount",
            "voucher": "AP Voucher Number",
            "type": "Type"
        } # Column map for standardized writing. Can change values for dynamic column names

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
            print("Successfully loaded transactional data from valid sheets.")

        except Exception as e:
            print("Error loading transactional detail file:", e)

    # Internal helper function that categorizes a row in CTIES file as Actual, Accrual, Reversal, etc.
    # Use by calling df["Type"] = df.apply(self._categorize_row, axis=1)
    def _categorize_row(self, row):
        '''
        Returns value for row 'Type' as a string
        '''
        voucher_number = str(row[self.colmap["voucher"]])
        amount = row[self.colmap["amount"]]

        if voucher_number.startswith("5"):
            return "Actual"
        elif voucher_number.startswith("2"):
            if amount > 0:
                return "Accrual"
            elif amount < 0:
                return "Reversal"
        elif voucher_number.startswith("9"):
            return "Reclass"
        else:
            return "Undefined"

    def get_transactional_data(self):
        '''
        Method that gets transactional data from C-TIES file. 
        Extracts all rows, categorizes them, and a returns a dict w/ actuals and accruals:

        Returns:
            dict: {
                'PO12345': {
                    'Jan': {
                        'Actual': 900
                        'Accrual': 950.0,
                        'Accrual Reversal': 0.0,
                    },
                    'Feb': {
                        'Actual': 800
                        'Accrual': 1050.0,
                        'Accrual Reversal': -950,
                    }
                    ...
                }
                ...
            }
        '''
        # Load and preprocess (categorization happens during load)
        self.load_transactional_detail_file()

        # Filter to only the columns needed, using instance variables
        cols = [
            self.colmap["po"],
            self.colmap["month"],
            self.colmap["amount"],
            self.colmap["type"]
        ]

        missing = [c for c in cols if c not in self.data.columns]
        if missing:
            raise ValueError(f"Missing required columns: {missing}") # Check for missing cols

        data_copy = self.data[cols].copy()

        # Ensure amount column is numeric
        data_copy[self.colmap["amount"]] = pd.to_numeric(
            data_copy[self.colmap["amount"]],
            errors='coerce'
        )

        # Filter valid transaction types using instance variable self.valid_types
        data_copy = data_copy[data_copy[self.colmap["type"]].isin(self.valid_types)]

        # Group by PO / Month / Type
        grouped = (
            data_copy.groupby([
                self.colmap["po"],
                self.colmap["month"],
                self.colmap["type"]
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

            # Actual values belong to prior month
            if month_num == 1:
                actual_month = "Dec"
            else:
                actual_month = self.month_map.get(month_num - 1)

            # Accruals/reversals belong to current month
            accrual_month = self.month_map.get(month_num)

            # Initialize PO
            if po not in result:
                result[po] = {}

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

        # Sort months for readability
        month_order = [
            "Jan","Feb","Mar","Apr","May","Jun",
            "Jul","Aug","Sep","Oct","Nov","Dec"
        ]

        for po in result:
            result[po] = {
                month: result[po][month]
                for month in month_order
                if month in result[po]
            }

        return result