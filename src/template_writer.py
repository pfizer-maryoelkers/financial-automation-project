from abc import ABC, abstractmethod
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

class TemplateWriterBase(ABC):
    """
    Abstract base class for writing data into Excel templates.
    Defines the interface for all template writers.
    """

    def __init__(self, template_path: str, sheet_name: str = None):
        """
        Initialize the template writer.
        :param template_path: Path to the Excel template file.
        :param sheet_name: Optional sheet name to work on.
        """
        self.wb = load_workbook(template_path)
        self.sheet = self.wb[sheet_name] if sheet_name else self.wb.active
        self.pos = []

    @abstractmethod
    def parse_template(self):
        """Parse template structure (headers, column mappings, etc.)."""
        pass

    @abstractmethod
    def write_data(self, data: dict):
        """
        Write data into the template.
        :param data: Dictionary of PO data structured by month.
        """
        pass

    def save(self, output_path: str = "template_output.xlsx"):
        """
        Save the workbook to the specified path.
        :param output_path: Output file path.
        """
        self.wb.save(output_path)


class FinancialTemplateV2Writer(TemplateWriterBase):
    """
    Writer for the financial template.
    Implements logic for writing PO data into the predefined template structure.
    """

    # Hardcoded column mapping for months
    COLUMN_MAP = {
        'Jan': {'Accrual Reversal': 'J', 'Forecast': 'K', 'Accrual': 'L', 'Actual': 'M'},
        'Feb': {'Accrual Reversal': 'O', 'Forecast': 'P', 'Accrual': 'Q', 'Actual': 'R'},
        'Mar': {'Accrual Reversal': 'T', 'Forecast': 'U', 'Accrual': 'V', 'Actual': 'W'},
        'Apr': {'Accrual Reversal': 'Y', 'Forecast': 'Z', 'Accrual': 'AA', 'Actual': 'AB'},
        'May': {'Accrual Reversal': 'AD', 'Forecast': 'AE', 'Accrual': 'AF', 'Actual': 'AG'},
        'Jun': {'Accrual Reversal': 'AI', 'Forecast': 'AJ', 'Accrual': 'AK', 'Actual': 'AL'},
        'Jul': {'Accrual Reversal': 'AN', 'Forecast': 'AO', 'Accrual': 'AP', 'Actual': 'AQ'},
        'Aug': {'Accrual Reversal': 'AS', 'Forecast': 'AT', 'Accrual': 'AU', 'Actual': 'AV'},
        'Sep': {'Accrual Reversal': 'AX', 'Forecast': 'AY', 'Accrual': 'AZ', 'Actual': 'BA'},
        'Oct': {'Accrual Reversal': 'BC', 'Forecast': 'BD', 'Accrual': 'BE', 'Actual': 'BF'},
        'Nov': {'Accrual Reversal': 'BH', 'Forecast': 'BI', 'Accrual': 'BJ', 'Actual': 'BK'},
        'Dec': {'Accrual Reversal': 'BM', 'Forecast': 'BN', 'Accrual': 'BO', 'Actual': 'BP'}
    }

    HEADER_ROW = 14
    PO_COLUMN = 'B'
    START_ROW = 15
    TEMPLATE_ROW = 20

    def parse_template(self):
        """
        Build a lookup of existing POs for update behavior.
        """
        self.po_to_row = {}
        for row in range(self.START_ROW, self.sheet.max_row + 1):
            val = self.sheet[f"{self.PO_COLUMN}{row}"].value
            if val and str(val).strip():
                self.po_to_row[str(val).strip()] = row
        
        # Create list of POs that we'll be using 
        self.pos = list(self.po_to_row.keys())

    def find_po_row(self, po: str) -> int:
        """Return row for existing PO or None."""
        return self.po_to_row.get(str(po).strip())

    def find_next_blank(self) -> int:
        """Find next blank row starting from START_ROW up to max_row."""
        for row in range(self.START_ROW, self.sheet.max_row + 1):
            if not self.sheet[f"{self.PO_COLUMN}{row}"].value:
                return row
        return None
    
    def insert_new_rows(self, num_rows: int):
        """
        Insert multiple rows below row 20.
        Each new row copies formulas from row 20 and clears non-formula cells.
        """
        source_row = 20  # Template row
        insert_at = source_row + 1  # Insert below template
        self.sheet.insert_rows(insert_at, amount=num_rows)

        # For each inserted row, copy formulas and formatting from row 20
        for offset in range(num_rows):
            new_row = insert_at + offset
            for src_cell, dest_cell in zip(self.sheet[source_row], self.sheet[new_row]):
                if src_cell.has_style:
                    dest_cell._style = src_cell._style
                dest_cell.number_format = src_cell.number_format
                if src_cell.data_type == 'f':  # formula
                    translator = Translator(src_cell.value, origin=src_cell.coordinate)
                    dest_cell.value = translator.translate_formula(dest_cell.coordinate)
                else:
                    dest_cell.value = None  # Keep new row blank except formulas

    def write_data(self, data: dict):
        """
        Pre-allocate rows if needed, then write data.
        """
        num_pos = len(data)
        available_rows = 20 - self.START_ROW  # Rows 15–19
        extra_rows = max(0, num_pos - available_rows)

        # Pre-allocate extra rows if needed
        if extra_rows > 0:
            self.insert_new_rows(extra_rows)

        # Write data into blank rows starting at row 15
        for po, months in data.items():
            row = self.find_po_row(po)
            if row is None:
                row = self.find_next_blank()
                self.sheet[f"{self.PO_COLUMN}{row}"] = po
                self.po_to_row[str(po).strip()] = row

            for month, values in months.items():
                if month not in self.COLUMN_MAP:
                    continue
                for key, col in self.COLUMN_MAP[month].items():
                    self.sheet[f"{col}{row}"] = values.get(key, 0)

        # Adding PO auto filter
        last_row = 20 + extra_rows
        self.sheet.auto_filter.ref = f"{self.PO_COLUMN}{self.HEADER_ROW}:{self.PO_COLUMN}{last_row}"

    def write_forecast_audit(self, forecast_df, selected_pos):
        """
        Write a forecast audit sheet containing only the POs
        that are being written to the main template.

        Parameters
        ----------
        forecast_df : pandas.DataFrame
            The raw forecast data loaded by ForecastReader.

        selected_pos : list[str]
            List of PO numbers that should appear in the audit sheet.
        """

        # 1. Filter DF to selected POs 
        if "PO #" not in forecast_df.columns:
            raise KeyError("Expected 'PO #' column not found in forecast dataframe.")

        forecast_df["PO #"] = forecast_df["PO #"].astype(str)
        audit_df = forecast_df[forecast_df["PO #"].isin(selected_pos)]

        audit_columns = [
            "PO #",
            "Pfizer requested PO",
            "IBM PMA",
            "PO Owner",
            "Work Code",
            "Project Name",
            "Project Start Date",
            "Project End Date",
            "Resource Talent Id",
            "Resource Name",
            "Location",
            "Resource Status",
            "Bill Rate",
            "Forecast Month Hours",
            "Forecast Fee $",
            "Forecast Expenses $",
            "Forecast Total $",
            "Actual Month Hours",
            "Actual Fee $",
            "Actual Expenses $",
            "Actual Total $",
            "Monthly Hours Variance",
            "Hours Billed Flag",
            "Forecast Accuracy",
            "Expense%",
            "Comments",
            "Total Forecast Hours 2025",
            "Jan 2025 - FTotal",
            "Feb 2025 - FTotal",
            "March 2025 - FTotal",
            "April 2025 - FTotal",
            "May 2025 - FTotal",
            "June 2025 - FTotal",
            "July 2025 - FTotal",
            "Aug 2025 - FTotal",
            "Sep 2025 - FTotal",
            "Oct 2025 - FTotal",
            "Nov 2025 - FTotal",
            "Dec 2025 - FTotal",
            "2025 Forecast total Fee"
        ]

        # Filter DF to only those columns (ignore missing columns gracefully)
        audit_df = audit_df[[c for c in audit_columns if c in audit_df.columns]]

        # 2. Create the audit sheet
        ws = self.wb.create_sheet("Forecast Source Data")

        # 3. Write header 
        for col_idx, column_name in enumerate(audit_df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=column_name)

        # 4. Write data rows
        for row_idx, (_, row) in enumerate(audit_df.iterrows(), start=2):
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        # Adding summation rows
        data_start_row = 2
        data_end_row = ws.max_row

        # Add total row ONE below the last data row
        total_row = data_end_row + 1

        # Label column (A)
        ws.cell(row=total_row, column=1, value="PO Total")

        # Add SUM formulas for numeric columns
        for col_idx, col_name in enumerate(audit_df.columns, start=1):
            # Skip non-numeric and non-summation columns
            if col_name not in [
                "Forecast Fee $",
                "Forecast Expenses $",
                "Forecast Total $",
                "Total Forecast Hours 2025",
                "Jan 2025 - FTotal",
                "Feb 2025 - FTotal",
                "March 2025 - FTotal",
                "April 2025 - FTotal",
                "May 2025 - FTotal",
                "June 2025 - FTotal",
                "July 2025 - FTotal",
                "Aug 2025 - FTotal",
                "Sep 2025 - FTotal",
                "Oct 2025 - FTotal",
                "Nov 2025 - FTotal",
                "Dec 2025 - FTotal",
                "2025 Forecast total Fee"
            ]:
                continue

            col_letter = get_column_letter(col_idx)
            formula = f"=SUBTOTAL(9, {col_letter}{data_start_row}:{col_letter}{data_end_row})"
            ws.cell(row=total_row, column=col_idx, value=formula)

        # 5. Add auto-filter + freeze header
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"

        # Column Grouping 
        # Group all columns between "PO #" and "Total Forecast Hours 2025"

        # Find column indexes
        columns = list(audit_df.columns)

        po_col_index = columns.index("PO #") + 1  # +1 because Excel is 1-based
        summary_col_index = columns.index("Total Forecast Hours 2025") + 1

        # --- Auto-adjust column widths ---
        for col_idx, col_name in enumerate(audit_df.columns, start=1):
            column_letter = get_column_letter(col_idx)
            max_length = len(str(col_name))

            # compute max length of all cell contents in this column
            for row_idx in range(2, len(audit_df) + 2):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None:
                    max_length = max(max_length, len(str(cell_value)))

            ws.column_dimensions[column_letter].width = max_length + 2

        # Group everything between PO # and summary column (exclusive)
        start_col = po_col_index + 1
        end_col = summary_col_index - 1

        if start_col <= end_col:
            ws.column_dimensions.group(
                get_column_letter(start_col),
                get_column_letter(end_col),
                hidden=True
            )

    # Helper function to categorize rows in transactional detail file
    def categorize_row(self, row):
        '''
        Returns value for row 'Type' as a string

        '''
        voucher_number = str(row['AP Voucher Number'])
        amount = row["GL Transaction Amount"]

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
        
         

    def write_transactions_audit(self, transactions_df, selected_pos):
        """
        Write a lightweight, trimmed, grouped Transactions Audit sheet.

        This refactored version simplifies the logic:
        - Only two groups: visible columns and non-visible columns.
        - Non-visible columns are grouped into one expandable block.
        """

        # --- 1. Normalize PO column ---
        if "PO Number" not in transactions_df.columns:
            raise KeyError("Expected 'PO Number' column not found in transactional dataframe.")

        transactions_df["PO Number"] = transactions_df["PO Number"].astype(str)

        # --- 2. Filter by selected POs ---
        audit_df = transactions_df[transactions_df["PO Number"].isin(selected_pos)].copy()

        # --- 3. Compute Type column ---
        # audit_df["Type"] = audit_df.apply(
        #     lambda row: (
        #         "Actual"
        #         if str(row["AP Voucher Number"]).startswith("5")
        #         else (
        #             "Accrual"
        #             if str(row["AP Voucher Number"]).startswith("2")
        #             and row["GL Transaction Amount"] > 0
        #             else (
        #                 "Reversal"
        #                 if str(row["AP Voucher Number"]).startswith("2")
        #                 and row["GL Transaction Amount"] < 0
        #                 else "Undefined"
        #             )
        #         )
        #     ),
        #     axis=1
        # )

        audit_df["Type"] = audit_df.apply(self.categorize_row, axis=1)

        # --- 4. Define visible columns ---
        visible_cols = [
            "PO Number",
            "Accounting Period",
            "AP Voucher Number",
            "Vendor Name",
            "WBS Element",
            "GL Invoice Date",
            "GL Posting Date",
            "GL Line Description",
            "Description",
            "Month",
            "GL Transaction Amount",
            "Type",
        ]

        # --- 5. Determine all non-visible columns automatically ---
        # Preserve order as it appears in audit_df
        non_visible = [c for c in audit_df.columns if c not in visible_cols]

        # --- 6. Final column ordering ---
        final_columns = visible_cols + non_visible
        audit_df = audit_df[final_columns]

        # --- 7. Create sheet ---
        ws = self.wb.create_sheet("Transactions Source Data")

        # --- 8. Write header ---
        for col_idx, col_name in enumerate(audit_df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=col_name)

        # --- 9. Write data rows ---
        for row_idx, (_, row) in enumerate(audit_df.iterrows(), start=2):
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        # --- 10. Build column index map ---
        col_to_idx = {name: idx + 1 for idx, name in enumerate(audit_df.columns)}

        # Apply filter to following columns: PO Number, Accounting Period, Type
        filter_cols = ["PO Number", "Accounting Period", "Type"]
        max_filter_col = max(col_to_idx[c] for c in filter_cols)
        last_filter_letter = get_column_letter(max_filter_col)

        ws.auto_filter.ref = f"A1:{last_filter_letter}1"

        # --- 11. Auto-adjust column widths ---
        for col_idx, col_name in enumerate(audit_df.columns, start=1):
            column_letter = get_column_letter(col_idx)
            max_length = len(str(col_name))

            # compute max length of all cell contents in this column
            for row_idx in range(2, len(audit_df) + 2):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None:
                    max_length = max(max_length, len(str(cell_value)))

            ws.column_dimensions[column_letter].width = max_length + 2

        
        # --- 12. Group the non-visible columns in a single block ---
        if non_visible:
            start = min(col_to_idx[col] for col in non_visible)
            end   = max(col_to_idx[col] for col in non_visible)
            ws.column_dimensions.group(
                get_column_letter(start),
                get_column_letter(end),
                hidden=True
            )