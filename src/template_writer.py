from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

class TemplateWriter:
    """Reads a template and builds a header map, then writes forecast data."""

    def __init__(self, template_path: str, sheet_name: str = "Template"):
        self.template_path = template_path
        self.sheet_name = sheet_name
        self.workbook = None
        self.sheet = None
        self.header_map = None

    def load_template(self):
        """Load the workbook and select the desired sheet."""
        self.workbook = load_workbook(self.template_path)
        self.sheet = self.workbook[self.sheet_name]

    def parse_headers(self) -> dict:
        """
        Parse header row (row 14) and build a mapping of field names and months to columns.
        Also applies an AutoFilter so headers are filterable.
        """
        if self.workbook is None or self.sheet is None:
            self.load_template()

        header_row = 14
        headers = [cell.value for cell in self.sheet[header_row]]

        categories = ["Forecast", "Actual", "Accrual Reversal", "Accrual"]
        self.header_map = {cat: {} for cat in categories}

        for col_idx, value in enumerate(headers, start=1):
            if not value:
                continue
            text = str(value)
            for cat in categories:
                if text.startswith(cat):
                    tokens = text.split()
                    if len(tokens) >= 2:
                        month = tokens[-1]
                        m = month[:3]  # Normalize to first 3 letters
                        self.header_map[cat][m] = col_idx

        # Apply AutoFilter from column B to last header column
        last_col = max(idx for idx, val in enumerate(headers, start=1) if val)
        first_filter_col = get_column_letter(2)  # B
        last_filter_col = get_column_letter(last_col)
        self.sheet.auto_filter.ref = f"{first_filter_col}{header_row}:{last_filter_col}{self.sheet.max_row}"

        return self.header_map

    def _find_first_blank_row(self) -> int:
        """Find first blank row in column B starting at row 15."""
        if self.workbook is None or self.sheet is None:
            self.load_template()
        row = 15
        while True:
            val = self.sheet.cell(row=row, column=2).value
            if val is None or str(val).strip() == "":
                return row
            row += 1

    def write_forecast(self, data: dict, po_filter: list[str] = None):
        """
        Write forecast values into the template for each PO.

        :param data: dict of PO -> month -> {'Forecast': value}
        :param po_filter: optional list of PO numbers to include; if None, all POs are written
        """
        if self.workbook is None or self.sheet is None:
            self.load_template()
        if self.header_map is None:
            self.parse_headers()

        row = self._find_first_blank_row()
        for po, month_dict in data.items():
            # Skip if filter is provided and PO is not in it
            if po_filter is not None and po not in po_filter:
                continue

            # Write PO into column B
            self.sheet.cell(row=row, column=2, value=po)

            # Write forecast values for each month
            for month, fields in month_dict.items():
                m = str(month)[:3]  # Normalize month
                col = self.header_map.get('Forecast', {}).get(m)
                if col:
                    val = fields.get('Forecast', None)
                    if val is not None:
                        self.sheet.cell(row=row, column=col, value=val)
            row += 1


    def save(self, out_path: str = None):
        """Save the workbook to the given path or overwrite the original."""
        if out_path is None:
            out_path = self.template_path
        self.workbook.save(out_path)
