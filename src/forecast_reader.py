
import pandas as pd

class ForecastReader:
    """Reads vendor forecast file and extracts forecast data."""

    def __init__(self, file_paths: list, po_col):
        """Initialize with the forecast file path."""
        self.file_paths = file_paths
        self.data = None
        self.po_col = po_col

    def _load_valid_sheet(self, file_path: str):
        """
        Attempts to read the first sheet that contains:
        - A 'self.po_col' column
        - At least one forecast column ending in '- FTotal'

        Returns:
            DataFrame if a valid sheet is found, otherwise None.
        """

        try:
            xls = pd.ExcelFile(file_path)
        except Exception as e:
            print(f"Error opening file {file_path}: {e}")
            return None

        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sheet)
            except Exception:
                continue

            # Verify required columns
            if self.po_col in df.columns:
                if any(col.endswith("- FTotal") for col in df.columns):
                    print(f"Selected sheet '{sheet}' from file {file_path}")
                    return df  # Found the correct sheet

        print(f"No valid sheet found in file {file_path}.")
        return None


    def load_forecast(self):
        """
        Load the forecast Excel file(s) into memory.
        Assumes self.file_path is always a list of one or more file paths.
        """

        dfs = []
        seen = set()           # Track POs already encountered
        dup_pos_total = set()  # Track POs appearing in later files

        for f in self.file_paths:
            try:
                df = self._load_valid_sheet(f)
                if df is None:
                    continue
                # Clean PO # column immediately after load
                df[self.po_col] = df[self.po_col].apply(
                    lambda x: str(int(float(x))) if str(x).replace('.', '', 1).isdigit() else str(x)
                )

            except Exception as e:
                print(f"Error reading file {f}: {e}")
                continue

            # Ensure the PO column exists
            if self.po_col not in df.columns:
                print(f"File {f} is missing the self.po_col column.")
                continue

            # Identify POs in this file
            pos = set(df[self.po_col].astype(str))


            # Detect duplicates across files
            intersection = seen.intersection(pos)
            if intersection:
                dup_pos_total |= intersection
                # Drop rows from later files for any duplicated PO
                df = df[~df[self.po_col].isin(intersection)]

            seen |= pos
            dfs.append(df)

        # Combine all valid DataFrames
        if dfs:
            self.data = pd.concat(dfs, ignore_index=True)
        else:
            self.data = None

        # Notify user of duplicate POs
        if dup_pos_total:
            print(
                f"\nWARNING: PO(s) {', '.join(map(str, sorted(dup_pos_total)))} "
                "appear in multiple forecast files. Only the first occurrence was used."
            )


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
        grouped = self.data.groupby(self.po_col)[forecast_cols].sum()

        # Build the model
        for po, row in grouped.iterrows():
            po_dict = {}
            po_df = self.data[self.data[self.po_col] == po]  # Original rows for this PO
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
   
