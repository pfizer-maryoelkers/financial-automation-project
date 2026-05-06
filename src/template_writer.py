from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font
import pandas as pd
from src.models import CostCenter, WBSCode, PO, MonthlyMetrics, ExceptionLog, ExceptionType


def month_sort_key(month_str):
    """Convert month string to sortable key for proper chronological ordering"""
    month_order = {
        'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
        'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12,
        'Unknown': 13
    }
    return month_order.get(month_str, 13)



class TemplateWriter:

    def __init__(self, 
                 file_path,
                 output_path,
                 overwrite,
                 header_row, 
                 po_column,
                 dec_acc_reversal_col, 
                 forecast_source_cols,
                 transactional_source_cols
        ):
        
        self.wb = load_workbook(file_path)
        self.sheet: Worksheet = self.wb.active  # type: ignore[assignment]
        if self.sheet is None:
            raise ValueError(f"Could not load active sheet from {file_path}")

        self.output_path = output_path

        # Overwrites previous months data if true
        self.overwrite = overwrite

        # Configs for template
        self.header_row = header_row    # Header row
        self.po_column = po_column    # Col where POs are entered


        # Cost centers, WBS codes, PO numbers, and their associated rows
        self.cost_centers = {}
        self.wbs_codes = {}

        # Column map (dynamically created starting with Dec Accrual Reversal)
        self.dec_acc_reversal_col = dec_acc_reversal_col # Col where first data entry exists
        self.column_map = self.get_column_map(starting_col=self.dec_acc_reversal_col)

        # Source sheet params
        self.forecast_source_cols = forecast_source_cols
        self.transactional_source_cols = transactional_source_cols

        # PO columns
        self.forecast_po_col = self.forecast_source_cols[0]
        self.transactional_po_col = self.transactional_source_cols[0]

    def get_column_map(self, starting_col):
        months = [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ]

        metrics = ["Accrual Reversal", "Forecast", "Accrual", "Actual"]

        col_map = {}
        col_index = column_index_from_string(starting_col)

        for month in months:
            month_map = {}
            for metric in metrics:
                month_map[metric] = get_column_letter(col_index)
                col_index += 1

            col_map[month] = month_map

            # Skip one column before next month's block
            col_index += 1

        return col_map

    def write_hierarchy(self, hierarchy: dict, pos: dict[str, int]):
        '''
        Writes hierarchy to template.
        Iterates CostCenter -> WBSCode -> PO and writes MonthlyMetrics to correct cells.
        Only writes to blank cells unless overwrite=True.
        '''
        for cc_id, cost_center in hierarchy.items():
            for wbs_code, wbs in cost_center.wbs_codes.items():
                for po_number, po in wbs.pos.items():
                    # Skip if PO not in template
                    if po_number not in pos:
                        print(f"PO '{po_number}' not found in template. Skipping.")
                        continue
                    row = pos[po_number]
                    for month, metrics in po.monthly_data.items():
                        if month not in self.column_map:
                            continue
                        month_cols = self.column_map[month]
                        values = {
                            'Accrual Reversal': metrics.accrual_reversal,
                            'Forecast': metrics.forecast,
                            'Accrual': metrics.accrual,
                            'Actual': metrics.actual
                        }
                        for metric, col_letter in month_cols.items():
                            cell = self.sheet[f"{col_letter}{row}"]
                            if self.overwrite or cell.value is None or str(cell.value).strip() == "":
                                cell.value = values.get(metric)



    ## Methods to write source sheets
    def write_forecast_source_sheet(self, forecast_df, pos: dict[str, int]):
        # Method to write forecast source sheet. 
        # Filter to template POs
        if self.forecast_po_col not in forecast_df.columns:
            raise KeyError(f"Expected {self.forecast_po_col} column not found in forecast dataframe.")

        forecast_df[self.forecast_po_col] = (
            forecast_df[self.forecast_po_col]
            .apply(lambda x: str(int(float(x))) if str(x).replace('.','',1).isdigit() else str(x))
        )
        filtered_df = forecast_df[forecast_df[self.forecast_po_col].isin(pos.keys())]

        visible_cols = [c for c in self.forecast_source_cols if c in filtered_df.columns]
        hidden_cols = [c for c in filtered_df.columns if c not in visible_cols]

        ordered_cols = visible_cols + hidden_cols
        source_df = filtered_df[ordered_cols]

        ws = self.wb.create_sheet("Forecast Source Data")

        for col_idx, col_name in enumerate(source_df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=col_name)

        for row_idx, (_, row) in enumerate(source_df.iterrows(), start=2):
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        data_start = 2
        data_end = ws.max_row
        total_row = data_end + 1
        ws.cell(row=total_row, column=1, value="PO Total")

        for col_idx, col_name in enumerate(source_df.columns, start=1):
            if col_name in visible_cols and col_name != self.forecast_po_col:
                letter = get_column_letter(col_idx)
                formula = f"=SUBTOTAL(9,{letter}{data_start}:{letter}{data_end})"
                ws.cell(row=total_row, column=col_idx, value=formula)

        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"

        for col_idx, col_name in enumerate(source_df.columns, start=1):
            letter = get_column_letter(col_idx)
            max_len = len(str(col_name))
            for row_idx in range(2, len(source_df) + 2):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val is not None:
                    max_len = max(max_len, len(str(val)))
            ws.column_dimensions[letter].width = max_len + 2

        if hidden_cols:
            start_idx = len(visible_cols) + 1
            end_idx = len(visible_cols) + len(hidden_cols)
            ws.column_dimensions.group(
                get_column_letter(start_idx),
                get_column_letter(end_idx),
                hidden=True
            )

    def write_transactional_source_sheet(self, transactions_df, pos: dict[str, int]):
        # Method to write transactional detail source sheet
        # Filter to POs present in the template
        if self.transactional_po_col not in transactions_df.columns:
            raise KeyError(f"Expected {self.transactional_po_col} column not found in transactional dataframe.")

        transactions_df[self.transactional_po_col] = transactions_df[self.transactional_po_col].astype(str)
        source_df = transactions_df[transactions_df[self.transactional_po_col].isin(pos.keys())]

        visible_cols = [c for c in self.transactional_source_cols if c in source_df.columns]
        hidden_cols = [c for c in source_df.columns if c not in visible_cols]

        final_cols = visible_cols + hidden_cols
        source_df = source_df[final_cols]

        ws = self.wb.create_sheet("Transactions Source Data")

        for col_idx, col_name in enumerate(source_df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=col_name)

        for row_idx, (_, row) in enumerate(source_df.iterrows(), start=2):
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"

        for col_idx, col_name in enumerate(source_df.columns, start=1):
            letter = get_column_letter(col_idx)
            max_len = len(str(col_name))
            for row_idx in range(2, len(source_df) + 2):
                value = ws.cell(row=row_idx, column=col_idx).value
                if value is not None:
                    max_len = max(max_len, len(str(value)))
            ws.column_dimensions[letter].width = max_len + 2

        if hidden_cols:
            start_idx = len(visible_cols) + 1
            end_idx = len(visible_cols) + len(hidden_cols)
            ws.column_dimensions.group(
                get_column_letter(start_idx),
                get_column_letter(end_idx),
                hidden=True
            )


    
    def write_exception_sheet(self, exception_log, transactional_df):
        ws = self.wb.create_sheet("Exceptions")
        
        # Define visible columns (Month moved to position 4 for better visibility)
        visible_headers = [
            'Cost Center', 'Month', 'WBS', 'PO',
            'Exception Type', 'Source Row', 'Amount', 'Type'
        ]
        
        # Get all transactional columns for hidden section
        # Exclude columns already shown in visible section
        excluded_cols = {'Cost Center*', 'WBS Element', 'PO Number', 'Month', 'GL BER Corp Amount', 'Type'}
        hidden_headers = [col for col in transactional_df.columns if col not in excluded_cols]
        
        all_headers = visible_headers + hidden_headers
        
        # Write headers
        for col_idx, header in enumerate(all_headers, start=1):
            ws.cell(row=1, column=col_idx, value=header)
        
        # Write data rows
        for row_idx, entry in enumerate(exception_log.entries, start=2):
            # Visible columns (updated order with Month at position 4)
            ws.cell(row=row_idx, column=1, value=entry.cost_center)
            ws.cell(row=row_idx, column=2, value=entry.month)
            ws.cell(row=row_idx, column=3, value=entry.wbs)
            ws.cell(row=row_idx, column=4, value=entry.po)
            ws.cell(row=row_idx, column=5, value=entry.exception_type.value)
            ws.cell(row=row_idx, column=6, value=entry.row_index)
            ws.cell(row=row_idx, column=7, value=entry.amount)
            ws.cell(row=row_idx, column=8, value=entry.transaction_type)
            
            # Hidden columns (full source row data)
            if entry.source_row_data:
                for col_idx_hidden, col_name in enumerate(hidden_headers, start=9):
                    ws.cell(row=row_idx, column=col_idx_hidden,
                           value=entry.source_row_data.get(col_name))
        
        # Apply formatting
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"
        
        # Auto-size visible columns
        for col_idx in range(1, len(visible_headers) + 1):
            letter = get_column_letter(col_idx)
            max_len = len(str(all_headers[col_idx - 1]))
            for row_idx in range(2, len(exception_log.entries) + 2):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None:
                    max_len = max(max_len, len(str(cell_value)))
            ws.column_dimensions[letter].width = min(max_len + 2, 50)
        
        # Group and hide supplementary columns
        if hidden_headers:
            start_idx = len(visible_headers) + 1
            end_idx = len(all_headers)
            ws.column_dimensions.group(
                get_column_letter(start_idx),
                get_column_letter(end_idx),
                hidden=True
            )

    def write_exception_data_sheet(self, exception_log):
        """Write raw exception data to hidden sheet for formula reference"""
        ws = self.wb.create_sheet("Exception_Data")
        
        # Headers
        headers = ['Cost Center', 'WBS', 'PO', 'Exception Type', 'Month', 'Amount', 'Type']
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_idx, value=header)
            ws.cell(row=1, column=col_idx).font = Font(bold=True)
        
        # Data rows
        for row_idx, entry in enumerate(exception_log.entries, start=2):
            ws.cell(row=row_idx, column=1, value=entry.cost_center or '')
            ws.cell(row=row_idx, column=2, value=entry.wbs or '')
            ws.cell(row=row_idx, column=3, value=entry.po or '')
            ws.cell(row=row_idx, column=4, value=entry.exception_type.value)
            ws.cell(row=row_idx, column=5, value=entry.month or '')
            ws.cell(row=row_idx, column=6, value=entry.amount)
            ws.cell(row=row_idx, column=7, value=entry.transaction_type or '')
        
        # Hide the sheet
        ws.sheet_state = 'hidden'
    
    def write_exception_summary_sheet(self, exception_log):
        """Create a summary sheet with interactive month filter showing exception counts by type and by cost center"""
        ws = self.wb.create_sheet("Exceptions Summary")
        
        # Get summary data for getting unique values
        summary_by_type = exception_log.summary_by_type()
        summary_by_cc = exception_log.summary_by_cost_center()
        
        # Get all unique exception types and cost centers
        all_exception_types = sorted(set(summary_by_type['counts'].keys()))
        all_cost_centers = sorted(set(summary_by_cc.keys()))
        
        # Get all unique months from exception data (ensure they are strings)
        all_months = sorted(set(str(entry.month) for entry in exception_log.entries if entry.month), key=month_sort_key)
        
        current_row = 1
        
        # Add Month Filter Dropdown
        ws.cell(row=current_row, column=1, value="Filter by Month:")
        ws.cell(row=current_row, column=1).font = Font(bold=True, size=12)
        
        # Create dropdown list (ensure all values are strings)
        month_options = ["All Months"] + [str(m) for m in all_months]
        dv = DataValidation(type="list", formula1=f'"{",".join(month_options)}"', allow_blank=False)
        dv.add(ws['B1'])
        ws.add_data_validation(dv)
        
        # Set default value
        ws['B1'] = "All Months"
        ws['B1'].font = Font(size=11)
        
        # Define named range for the filter cell
        from openpyxl.workbook.defined_name import DefinedName
        defined_name = DefinedName('MonthFilter', attr_text=f"'{ws.title}'!$B$1")
        self.wb.defined_names['MonthFilter'] = defined_name
        
        current_row += 2
        
        # Section 1: Summary by Exception Type (with dynamic formulas)
        ws.cell(row=current_row, column=1, value="Exceptions Summary by Type")
        ws.cell(row=current_row, column=1).font = Font(bold=True, size=14)
        current_row += 2
        
        # Headers for type summary
        ws.cell(row=current_row, column=1, value="Exception Type")
        ws.cell(row=current_row, column=2, value="Count")
        ws.cell(row=current_row, column=3, value="% of Total")
        for col in range(1, 4):
            ws.cell(row=current_row, column=col).font = Font(bold=True)
        current_row += 1
        
        # Data rows for type summary with formulas
        type_start_row = current_row
        for exc_type in all_exception_types:
            ws.cell(row=current_row, column=1, value=exc_type)
            
            # Count formula: IF MonthFilter="All Months", count all, else count for specific month
            count_formula = (
                f'=IF(MonthFilter="All Months",'
                f'COUNTIF(Exception_Data!$D:$D,"{exc_type}"),'
                f'COUNTIFS(Exception_Data!$D:$D,"{exc_type}",Exception_Data!$E:$E,MonthFilter))'
            )
            ws.cell(row=current_row, column=2, value=count_formula)
            
            # Percentage formula
            pct_formula = f'=IF(SUM($B${type_start_row}:$B${type_start_row + len(all_exception_types) - 1})=0,0,B{current_row}/SUM($B${type_start_row}:$B${type_start_row + len(all_exception_types) - 1})*100)'
            ws.cell(row=current_row, column=3, value=pct_formula)
            ws.cell(row=current_row, column=3).number_format = '0.0"%"'
            
            current_row += 1
        
        # Total row for type summary
        ws.cell(row=current_row, column=1, value="TOTAL")
        ws.cell(row=current_row, column=1).font = Font(bold=True)
        total_formula = f'=SUM(B{type_start_row}:B{current_row - 1})'
        ws.cell(row=current_row, column=2, value=total_formula)
        ws.cell(row=current_row, column=2).font = Font(bold=True)
        ws.cell(row=current_row, column=3, value='100.0%')
        ws.cell(row=current_row, column=3).font = Font(bold=True)
        current_row += 3
        
        # Section 2: Summary by Cost Center (with dynamic formulas)
        ws.cell(row=current_row, column=1, value="Exceptions Summary by Cost Center")
        ws.cell(row=current_row, column=1).font = Font(bold=True, size=14)
        current_row += 2
        
        # Headers for cost center summary
        ws.cell(row=current_row, column=1, value="Cost Center")
        ws.cell(row=current_row, column=2, value="Total")
        for idx, exc_type in enumerate(all_exception_types, start=3):
            ws.cell(row=current_row, column=idx, value=exc_type)
        for col in range(1, len(all_exception_types) + 3):
            ws.cell(row=current_row, column=col).font = Font(bold=True)
        current_row += 1
        
        # Data rows for cost center summary with formulas
        cc_start_row = current_row
        for cc in all_cost_centers:
            ws.cell(row=current_row, column=1, value=cc)
            
            # Total formula for this cost center
            total_formula = (
                f'=IF(MonthFilter="All Months",'
                f'COUNTIF(Exception_Data!$A:$A,"{cc}"),'
                f'COUNTIFS(Exception_Data!$A:$A,"{cc}",Exception_Data!$E:$E,MonthFilter))'
            )
            ws.cell(row=current_row, column=2, value=total_formula)
            
            # Count by exception type
            for idx, exc_type in enumerate(all_exception_types, start=3):
                type_formula = (
                    f'=IF(MonthFilter="All Months",'
                    f'COUNTIFS(Exception_Data!$A:$A,"{cc}",Exception_Data!$D:$D,"{exc_type}"),'
                    f'COUNTIFS(Exception_Data!$A:$A,"{cc}",Exception_Data!$D:$D,"{exc_type}",Exception_Data!$E:$E,MonthFilter))'
                )
                ws.cell(row=current_row, column=idx, value=type_formula)
            
            current_row += 1
        
        # Auto-size columns
        max_cols = max(len(all_exception_types) + 3, 10)
        for col_idx in range(1, max_cols):
            letter = get_column_letter(col_idx)
            max_len = 10
            for row_idx in range(1, current_row):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None and not str(cell_value).startswith('='):
                    max_len = max(max_len, len(str(cell_value)))
            ws.column_dimensions[letter].width = min(max_len + 2, 50)
        
        # Freeze panes at row 4 (below filter and title)
        ws.freeze_panes = "A4"
        
    
    def save(self):
        """Saves the workbook to the output path."""
        try:
            self.wb.save(self.output_path)
            print(f"Workbook saved to: {self.output_path}")
        except Exception as e:
            raise Exception(f"Failed to save workbook: {e}")
        
