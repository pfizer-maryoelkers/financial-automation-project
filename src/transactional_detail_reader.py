
import pandas as pd

class TransactionalDetailReader:
    """Parent class for reading transactional detail file and extracts accruals, actuals, and reversals.
    
    This class has the following subclasses:
        InvoiceActualReader
        AccrualReader
        ReclassReader

    Each of these subclasses includes the logic to parse the specific values, based on the following identifiers:
    
    Codes for AP Voucher Number:
        210 - Accrual/Reversal
        510 - Invoice
        900 - Reclass
    
    """

    def __init__(self, file_path: str, sheet_name="Transaction Detail Report"):
        """Initialize with the transactional detail file path."""
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.data = None

    def load_transactional_detail(self):
        """Load the transactional Excel file."""
        try:
            self.data = pd.read_excel(self.file_path, sheet_name=self.sheet_name, header=1)
        except Exception as e:
            print(e)

class InvoiceActualReader(TransactionalDetailReader):
    def get_transactional_data(self) -> dict:
        """
        Extract PO and monthly actuals.
        Returns:
            dict: {
                'PO12345': {
                    'Jan': {'Actual': 900, 'Source': [row_indices]},
                    'Feb': {...}
                }
            }
        """
        # 1. Ensure data is loaded
        if self.data is None:
            try:
                self.load_transactional_detail()
            except Exception:
                raise ValueError("Transactional detail data not loaded. Call load_transactional_detail() first.")

        # 2. Filter rows for actuals (GL prefix = 5)
        actuals_df = self.data.copy()

        # Ensure GL Invoice Number column exists
        if 'AP Voucher Number' not in actuals_df.columns:
            raise KeyError("'AP Voucher Number' column not found in transactional detail data.")

        # Filter rows where GL Invoice Number starts with '5'
        actuals_df = actuals_df[actuals_df['AP Voucher Number'].astype(str).str.startswith('5')]


        # 3. Normalize months and apply offset (+1 month for actuals)
        if 'Month' not in actuals_df.columns:
            raise KeyError("'Month' column not found in transactional detail data.")

        # Apply offset and wrap around (e.g., Dec + 1 → Jan)
        actuals_df['AdjustedMonth'] = (actuals_df['Month'] + 1 - 1) % 12 - 1

        # Map numeric month to abbreviation
        month_map = {
            1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
            7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
        }
        actuals_df['NormalizedMonth'] = actuals_df['AdjustedMonth'].map(month_map)


        # 4. Convert numeric values (GL Transaction Amount)
        if 'GL Transaction Amount' not in actuals_df.columns:
            raise KeyError("'GL Transaction Amount' column not found in transactional detail data.")

        actuals_df['GL Transaction Amount'] = pd.to_numeric(
            actuals_df['GL Transaction Amount'], errors='coerce'
        ).fillna(0.0)


        # 5. Aggregate by PO and month
        if 'PO Number' not in actuals_df.columns:
            raise KeyError("'PO Number' column not found in transactional detail data.")

        grouped = actuals_df.groupby(['PO Number', 'NormalizedMonth'])['GL Transaction Amount'].sum().reset_index()


        # 6. Build result dictionary with 'Actual' and 'Source'
        result = {}

        for _, row in grouped.iterrows():
            po = str(row['PO Number'])
            month = row['NormalizedMonth']
            value = row['GL Transaction Amount']

            # Find source rows in actuals_df for this PO and month
            source_rows = actuals_df[
                (actuals_df['PO Number'] == row['PO Number']) &
                (actuals_df['NormalizedMonth'] == month)
            ].index.tolist()

            # Initialize PO entry if not exists
            if po not in result:
                result[po] = {}

            # Add month data
            result[po][month] = {
                'Actual': float(value),
                'Source': source_rows
            }

        # Reordering months for easier debugging:
        month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        
        for po in result:
            # Sort the inner dict by month_order
            result[po] = {month: result[po][month] for month in month_order if month in result[po]}

        return result

class AccrualReader(TransactionalDetailReader):
    pass

class ReclassReader(TransactionalDetailReader):
    pass