
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
            dict: { 'PO12345': {'Jan': {'Forecast': 1000}, 'Feb': {...}} }
        """
        if self.data is None:
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
            for col in forecast_cols:
                month = month_map[col]
                value = row[col] if not pd.isna(row[col]) else 0
                po_dict[month] = {'Forecast': float(value)}
            result[str(po)] = po_dict

        return result
   
