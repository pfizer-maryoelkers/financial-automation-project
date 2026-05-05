from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, column_index_from_string



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
        self.sheet = self.wb.active

        self.output_path = output_path

        # Overwrites previous months data if true
        self.overwrite = overwrite

        # Configs for template
        self.header_row = header_row    # Header row
        self.po_column = po_column    # Col where POs are entered


        # Cost centers, WBS codes, PO numbers, and their associated rows
        self.cost_centers = {}
        self.wbs_codes = {}
        self.pos = self.get_existing_pos()

        # Column map (dynamically created starting with Dec Accrual Reversal)
        self.dec_acc_reversal_col = dec_acc_reversal_col # Col where first data entry exists
        self.column_map = self.get_column_map(starting_col=self.dec_acc_reversal_col)

        # Source sheet params
        self.forecast_source_cols = forecast_source_cols
        self.transactional_source_cols = transactional_source_cols

        # PO columns
        self.forecast_po_col = self.forecast_source_cols[0]
        self.transactional_po_col = self.transactional_source_cols[0]

    ## Methods to get existing cost centers, WBS codes, and POs from input template sheet
    def get_existing_cost_centers(self):
        # Gets existing cost centers (if exists)
        # TODO: implement
        pass

    def get_existing_wbs(self):
        # Gets existing WBS codes (if exists)
        # TODO: implement
        pass
    
    def get_existing_pos(self):
        # Gets existing POs in the template
        self.pos = {}
        row = self.header_row + 1 # Starting at row below header row

        # Find stop_row
        stop_row = None
        for search_row in range(1, self.sheet.max_row + 1):
            if self.sheet[f"A{search_row}"].value == "Previous Period Invoices":
                stop_row = search_row - 1
                break


        while row < stop_row:
            cell = self.sheet[f"{self.po_column}{row}"].value

            # Normalize PO value
            po = str(cell).strip()

            # Store mapping PO → row
            self.pos[po] = row

            row += 1
        
        return self.pos
    
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

    def write_data(self, data:dict):
        '''
        Writes data to template.

        Expects data dict in format:

            dict: {
                'PO12345': {
                    'Jan': {
                        'Forecast': 1050,
                        'Actual': 900
                        'Accrual': 950.0,
                        'Accrual Reversal': 0.0,
                    },
                    'Feb': {
                        'Forecast': 950,
                        'Actual': 800
                        'Accrual': 1050.0,
                        'Accrual Reversal': -950,
                    }
                    ...
                }
                ...
            }

        '''

        for po, row in self.pos.items():

            # If PO is not found in the data → leave row blank and print
            if po not in data:
                print(f"PO '{po}' found in template but not in source data. Leaving blank.")
                continue

            # Write PO into the sheet (if the blank template expects it)
            self.sheet[f"{self.po_column}{row}"] = po

            # Loop through months for this PO
            for month, metrics in data[po].items():
                if month not in self.column_map:
                    continue  # ignore unexpected months

                # Loop through metrics (Forecast, Actual, Accrual, Accrual Reversal)
                for metric, col_letter in self.column_map[month].items():
                    cell = self.sheet[f"{col_letter}{row}"]
                    # If overwrite is True, write to cells. Otherwise only write if cell is blank
                    if self.overwrite or cell.value is None or str(cell.value).strip() == "":
                        value = metrics.get(metric)
                        cell.value = value

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
    def write_forecast_source_sheet(self, forecast_df):
        # Method to write forecast source sheet. 
        # Filter to template POs
        if self.forecast_po_col not in forecast_df.columns:
            raise KeyError(f"Expected {self.forecast_po_col} column not found in forecast dataframe.")

        forecast_df[self.forecast_po_col] = (
            forecast_df[self.forecast_po_col]
            .apply(lambda x: str(int(float(x))) if str(x).replace('.','',1).isdigit() else str(x))
        )        
        filtered_df = forecast_df[forecast_df[self.forecast_po_col].isin(self.pos.keys())]

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

    def write_transactional_source_sheet(self, transactions_df):
        # Method to write transactional detail source sheet
        # Filter to POs present in the template
        if self.transactional_po_col not in transactions_df.columns:
            raise KeyError(f"Expected {self.transactional_po_col} column not found in transactional dataframe.")

        transactions_df[self.transactional_po_col] = transactions_df[self.transactional_po_col].astype(str)
        source_df = transactions_df[transactions_df[self.transactional_po_col].isin(self.pos.keys())]

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
        
        # Define visible columns
        visible_headers = [
            'Cost Center', 'WBS', 'PO', 'Exception Type',
            'Source Row', 'Month', 'Amount', 'Type'
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
            # Visible columns
            ws.cell(row=row_idx, column=1, value=entry.cost_center)
            ws.cell(row=row_idx, column=2, value=entry.wbs)
            ws.cell(row=row_idx, column=3, value=entry.po)
            ws.cell(row=row_idx, column=4, value=entry.exception_type.value)
            ws.cell(row=row_idx, column=5, value=entry.row_index)
            ws.cell(row=row_idx, column=6, value=entry.month)
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

    def write_exception_summary_sheet(self, exception_log):
        """Create a summary sheet showing exception counts by type and by cost center"""
        ws = self.wb.create_sheet("Exceptions Summary")
        
        # Get summary data
        summary_by_type = exception_log.summary_by_type()
        summary_by_cc = exception_log.summary_by_cost_center()
        
        current_row = 1
        
        # Section 1: Summary by Exception Type
        ws.cell(row=current_row, column=1, value="Exceptions Summary by Type")
        ws.cell(row=current_row, column=1).font = ws.cell(row=current_row, column=1).font.copy(bold=True, size=14)
        current_row += 2
        
        # Headers for type summary
        ws.cell(row=current_row, column=1, value="Exception Type")
        ws.cell(row=current_row, column=2, value="Count")
        ws.cell(row=current_row, column=3, value="% of Total")
        for col in range(1, 4):
            ws.cell(row=current_row, column=col).font = ws.cell(row=current_row, column=col).font.copy(bold=True)
        current_row += 1
        
        # Data rows for type summary
        type_start_row = current_row
        for exc_type, count in sorted(summary_by_type['counts'].items()):
            ws.cell(row=current_row, column=1, value=exc_type)
            ws.cell(row=current_row, column=2, value=count)
            percentage = summary_by_type['percentages'][exc_type]
            ws.cell(row=current_row, column=3, value=f"{percentage:.1f}%")
            current_row += 1
        
        # Total row for type summary
        ws.cell(row=current_row, column=1, value="TOTAL")
        ws.cell(row=current_row, column=2, value=summary_by_type['total'])
        ws.cell(row=current_row, column=3, value="100.0%")
        for col in range(1, 4):
            ws.cell(row=current_row, column=col).font = ws.cell(row=current_row, column=col).font.copy(bold=True)
        current_row += 3
        
        # Section 2: Summary by Cost Center
        ws.cell(row=current_row, column=1, value="Exceptions Summary by Cost Center")
        ws.cell(row=current_row, column=1).font = ws.cell(row=current_row, column=1).font.copy(bold=True, size=14)
        current_row += 2
        
        # Get all exception types for column headers
        all_exception_types = sorted(set(summary_by_type['counts'].keys()))
        
        # Headers for cost center summary
        ws.cell(row=current_row, column=1, value="Cost Center")
        ws.cell(row=current_row, column=2, value="Total")
        for idx, exc_type in enumerate(all_exception_types, start=3):
            ws.cell(row=current_row, column=idx, value=exc_type)
        for col in range(1, len(all_exception_types) + 3):
            ws.cell(row=current_row, column=col).font = ws.cell(row=current_row, column=col).font.copy(bold=True)
        current_row += 1
        
        # Data rows for cost center summary
        for cc, data in sorted(summary_by_cc.items()):
            ws.cell(row=current_row, column=1, value=cc)
            ws.cell(row=current_row, column=2, value=data['total'])
            for idx, exc_type in enumerate(all_exception_types, start=3):
                count = data['by_type'].get(exc_type, 0)
                ws.cell(row=current_row, column=idx, value=count if count > 0 else '')
            current_row += 1
        
        # Auto-size columns
        for col_idx in range(1, len(all_exception_types) + 3):
            letter = get_column_letter(col_idx)
            max_len = 10
            for row_idx in range(1, current_row):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None:
                    max_len = max(max_len, len(str(cell_value)))
            ws.column_dimensions[letter].width = min(max_len + 2, 50)
        
        # Freeze panes at row 2 for first section
        ws.freeze_panes = "A2"
        
    
    def save(self):
        """Saves the workbook to the output path."""
        try:
            self.wb.save(self.output_path)
            print(f"Workbook saved to: {self.output_path}")
        except Exception as e:
            raise Exception(f"Failed to save workbook: {e}")
        
