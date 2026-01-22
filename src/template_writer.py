from abc import ABC, abstractmethod
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

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