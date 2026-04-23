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
                 forecast_sum_exclude_cols,
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
        # self.forecast_sum_exclude_cols = forecast_sum_exclude_cols #TODO: remove
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

            # if cell is None or str(cell).strip() == "":
            #     break  # stop at the first blank cell

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


    ##NOTE: I think we can remove this helper function - but keeping here in case we need to add it back.
    
    ## Methods to write data to sheet
    # def _insert_new_rows(self, num_rows):
    #     """
    #     Helper function to insert new data entry rows.
    #     Since template only has 2 rows dedicated for data entry, this function is needed we have more than 2 POs to write.
    #     """
    #     source_row = self.header_row + 2  # Template row
    #     insert_at = source_row + 1  # Insert below template

    #     self.sheet.insert_rows(insert_at, amount=num_rows)

    #     # For each inserted row, copy formulas and formatting from source row
    #     for offset in range(num_rows):
    #         new_row = insert_at + offset
    #         for src_cell, dest_cell in zip(self.sheet[source_row], self.sheet[new_row]):
    #             if src_cell.has_style:
    #                 dest_cell._style = src_cell._style
    #             dest_cell.number_format = src_cell.number_format
    #             if src_cell.data_type == 'f':  # formula
    #                 translator = Translator(src_cell.value, origin=src_cell.coordinate)
    #                 dest_cell.value = translator.translate_formula(dest_cell.coordinate)
    #             else:
    #                 dest_cell.value = None  # Keep new row blank except formulas

    #     print(f"Inserted {num_rows} blank rows.\n")
        

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
        #NOTE: this logic relies on depracated _insert_blank_rows() method above. 

        # num_pos = len(self.pos)
        # extra_rows = max(0, num_pos - 2)

        # # Pre-allocate extra rows if needed
        # if extra_rows > 0:
        #     self._insert_new_rows(extra_rows)


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

        # print("======== DEBUG ========")
        # print("POs:")
        # print(self.pos.keys())
        # print("\n")
        # print("Forecast DF POs:")
        # print(forecast_df["PO #"])
        # print("\n Filtered DF:")
        # print(filtered_df)
        # print("======== END DEBUG ========")

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


    def save(self):
        """Saves the workbook to the output path."""
        try:
            self.wb.save(self.output_path)
            print(f"Workbook saved to: {self.output_path}")
        except Exception as e:
            raise Exception(f"Failed to save workbook: {e}")