
import pandas as pd

class ForecastReader:
    """Reads vendor forecast file and extracts forecast data."""

    def __init__(self, file_path: str, sheet_name="T&M Details"):
        """Initialize with the forecast file path."""
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.data = None

    def load_forecast(self):
        """Load the forecast Excel file into memory."""
        try:
            self.data = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
        except Exception as e:
            print(e)


    def get_forecast_data(self) -> dict:
        """
        Extract PO and monthly forecast values.
        Returns:
            dict: { 'PO12345': 
                {'Jan': 
                    {'Forecast': 1000, 'Source': source}, 
                'Feb': 
                    {'Forecast': 2000, 'Source': source}, 
                ...
                } 
            }
        """
        if self.data is None:
            try:
                self.load_forecast()
            except:
                raise ValueError("Forecast data not loaded. Call load_forecast() first.")

        # Identify forecast columns (those ending with '- FTotal')
        forecast_cols = [col for col in self.data.columns if col.endswith('- FTotal')]

        # Normalize month names (e.g., 'Jan', 'Feb', 'Mar')
        month_map = {}
        for col in forecast_cols:
            # Example: "Jan 2025 - FTotal" → "Jan"
            month_name = col.split()[0][:3]  # Take first 3 letters for consistency
            month_map[col] = month_name

        # Normalize numeric data
        for col in forecast_cols:
            self.data[col] = pd.to_numeric(self.data[col], errors="coerce").fillna(0.0)


        # Initialize result dictionary
        result = {}

        # Group by PO and sum forecast values across rows (multiple resources per PO)
        grouped = self.data.groupby('PO #')[forecast_cols].sum()

        # Build the model
        for po, row in grouped.iterrows():
            po_dict = {}
            po_df = self.data[self.data['PO #'] == po]  # Original rows for this PO
            for col in forecast_cols:
                # Getting month
                month = month_map[col]
                
                # Getting forecats value
                value = row[col] if not pd.isna(row[col]) else 0
                
                # Find source rows contributing to this forecast
                source_rows = po_df.loc[po_df[col].notna() & (po_df[col] != 0)].index.tolist()
                
                # Adding to po_dict
                po_dict[month] = {'Forecast': float(value), 'Source': source_rows}
            result[str(po)] = po_dict

        return result
   
