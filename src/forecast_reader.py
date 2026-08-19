
import pandas as pd

# Month abbreviation → full name prefix used in FTotal column names
_MONTH_ABBR = {
    1: "Jan", 2: "Feb", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "Aug",
    9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
}

class ForecastReader:
    """Reads vendor forecast file and extracts forecast data."""

    def __init__(self, file_paths: list, po_col):
        """Initialize with the forecast file path."""
        self.file_paths = file_paths
        self.data = None
        self.po_col = po_col

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _normalise_po_col(self, df: pd.DataFrame) -> pd.DataFrame:
        """Rename a bare 'PO#' column (no space) to self.po_col ('PO #') so
        downstream logic only needs to check one name."""
        alt = "PO#"
        if self.po_col not in df.columns and alt in df.columns:
            df = df.rename(columns={alt: self.po_col})
        return df

    def _billing_month_label(self, xls: pd.ExcelFile, sheet: str) -> str | None:
        """Try to read the billing-period month label from the sheet metadata rows.

        The IBM T&M forecast sheets store a header block in the first few rows
        before the data table:
          Row 1 col A: 'Billing Period Start Date'  col B: <start date>
          Row 2 col A: 'Billing Period End Date'    col B: <end date>

        We prefer the row labelled 'Billing Period End Date' because that
        determines the accounting period month (e.g. AP02 ends Feb 20 → "Feb").
        If no labelled row is found we fall back to the first date found.

        Returns a label like 'Feb 2026' or None if it cannot be determined.
        """
        try:
            meta = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=5)
        except Exception:
            return None

        fallback_date: pd.Timestamp | None = None

        for _, row in meta.iterrows():
            label = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""
            # Prefer the end-date row
            if "end date" in label:
                for cell in row[1:]:
                    if isinstance(cell, pd.Timestamp):
                        month_name = _MONTH_ABBR.get(cell.month)
                        if month_name:
                            return f"{month_name} {cell.year}"
            # Capture any date as fallback
            if fallback_date is None:
                for cell in row:
                    if isinstance(cell, pd.Timestamp):
                        fallback_date = cell
                        break

        if fallback_date is not None:
            month_name = _MONTH_ABBR.get(fallback_date.month)
            if month_name:
                return f"{month_name} {fallback_date.year}"

        return None

    def _synthesise_ftotal(self, df: pd.DataFrame, month_label: str) -> pd.DataFrame:
        """Add a '<Month> - FTotal' column by converting ForecastTotalFee.

        This handles the single-period IBM T&M sheet format where there is no
        per-month breakdown — the sheet covers exactly one billing period and
        'ForecastTotalFee' holds the total forecast for that period.

        Formula errors and non-numeric values are coerced to 0.
        """
        col_name = f"{month_label} - FTotal"
        raw = df["ForecastTotalFee"].apply(
            lambda v: pd.to_numeric(v, errors="coerce") if not isinstance(v, dict) else 0.0
        )
        df[col_name] = raw.fillna(0.0)
        return df

    def _forward_fill_po(self, df: pd.DataFrame) -> pd.DataFrame:
        """Forward-fill the PO column so that resource rows (where PO# is null)
        inherit the PO number from their parent row."""
        df[self.po_col] = df[self.po_col].replace(
            ["(blank)", "nan", "None"], pd.NA
        )
        df[self.po_col] = df[self.po_col].ffill()
        return df

    def _load_valid_sheet(self, file_path: str):
        """
        Attempts to read the first sheet that contains:
        - A 'self.po_col' (or bare 'PO#') column
        - At least one forecast column ending in '- FTotal',
          OR a 'ForecastTotalFee' column (single-period IBM T&M format)

        For the single-period format a synthetic '- FTotal' column is created
        from ForecastTotalFee using the billing month from the sheet header.

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

            df = self._normalise_po_col(df)

            if self.po_col not in df.columns:
                continue

            # Case 1: already has per-month FTotal columns
            if any(col.endswith("- FTotal") for col in df.columns):
                print(f"Selected sheet '{sheet}' from file {file_path}")
                df = self._forward_fill_po(df)
                return df

            # Case 2: single-period IBM T&M / FP sheet with ForecastTotalFee
            if "ForecastTotalFee" in df.columns:
                month_label = self._billing_month_label(xls, sheet)
                if month_label is None:
                    print(
                        f"WARNING: Could not determine billing month for sheet "
                        f"'{sheet}' in {file_path}. Skipping."
                    )
                    continue
                print(
                    f"Selected sheet '{sheet}' from file {file_path} "
                    f"(single-period format, mapped to '{month_label}')"
                )
                df = self._forward_fill_po(df)
                df = self._synthesise_ftotal(df, month_label)
                return df

        print(f"No valid sheet found in file {file_path}.")
        return None


    @staticmethod
    def _clean_po_value(x) -> str:
        """Normalise a single PO cell value to a clean string integer where possible."""
        s = str(x).strip()
        # Convert float-looking integers (e.g. 9500905777.0) to plain int string
        if s.replace('.', '', 1).replace('-', '', 1).isdigit():
            try:
                return str(int(float(s)))
            except (ValueError, OverflowError):
                pass
        return s

    def _expand_slash_pos(self, df: 'pd.DataFrame') -> 'pd.DataFrame':
        """Expand rows where PO # contains slash-separated values (e.g. '9500905777/9501153096').

        Each slash-separated PO gets its own row with the same forecast data, so
        both PO numbers receive the forecast when looked up individually.
        """
        import pandas as pd
        po_col = self.po_col
        rows = []
        for _, row in df.iterrows():
            raw = str(row[po_col]).strip()
            if '/' in raw:
                parts = [p.strip() for p in raw.split('/') if p.strip()]
                for part in parts:
                    new_row = row.copy()
                    new_row[po_col] = self._clean_po_value(part)
                    rows.append(new_row)
            else:
                rows.append(row)
        if rows:
            return pd.DataFrame(rows, columns=df.columns).reset_index(drop=True)
        return df.iloc[0:0]  # empty frame with same columns

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
                # Clean PO # column: normalise numeric strings, then expand slash-pairs
                df[self.po_col] = df[self.po_col].apply(self._clean_po_value)
                df = self._expand_slash_pos(df)

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
            except Exception as e:
                raise ValueError(
                    f"Forecast data was not loaded successfully, and automatic loading failed. "
                    f"Please verify that the forecast file path exists and is a valid Excel document. "
                    f"Internal load error: {e}"
                )

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
   
