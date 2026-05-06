from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

class TemplateReader:

    def __init__(self,
                file_path,
                header_row,
                po_col,
                po_stop_marker,
                cost_center_col,
                cost_center_start_row
        ):
        
        
        self.wb = load_workbook(file_path)
        self.sheet: Worksheet = self.wb.active  # type: ignore[assignment]
        
        # Ensure sheet was loaded successfully
        if self.sheet is None:
            raise ValueError(f"Could not load active sheet from {file_path}")

        # Initializing instance variable
        self.header_row = header_row
        self.po_col = po_col
        self.po_stop_marker = po_stop_marker
        self.cost_center_col = cost_center_col
        self.cost_center_start_row = cost_center_start_row

        # Read on init
        self.cost_centers = self.get_existing_cost_centers()
        self.pos = self.get_existing_pos()

    def get_existing_cost_centers(self) -> list[str]:
        """
        Reads cost centers from configured column starting at configured row.
        Stops at first blank cell.
        Returns:
            list[str]: e.g. ['1234', '2345', 'CC-999']
        """

        #NOTE: Possible extension: verify cost centers with unique cost centers existing in column J (or whatever column cost centers are in)

        cost_centers = []
        row = self.cost_center_start_row

        while True:
            cell = self.sheet[f"{self.cost_center_col}{row}"].value
            if cell is None or str(cell).strip() == "":
                break
            cost_centers.append(str(cell).strip())
            row += 1

        if not cost_centers:
            print("WARNING: No cost centers found in template.")
        else:
            print(f"Found {len(cost_centers)} cost centers: {cost_centers}")
        return cost_centers

    def get_existing_pos(self) -> dict[str, int]:
        """
        Extract PO numbers and their row positions from the template.
        
        Returns:
            dict[str, int]: Mapping of PO number to row number
            
        Raises:
            ValueError: If stop marker is not found in template
        """
        stop_row = self._find_stop_row()
        pos = self._extract_pos_from_rows(stop_row)
        self._log_pos_summary(pos)
        return pos
    
    def _find_stop_row(self) -> int:
        """Find the row containing the stop marker."""
        max_row = self.sheet.max_row or 1000  # Fallback to reasonable default
        for search_row in range(1, max_row + 1):
            if self.sheet[f"A{search_row}"].value == self.po_stop_marker:
                return search_row
        
        raise ValueError(
            f"Could not find '{self.po_stop_marker}' marker in template. "
            "Please check the template format or update po_stop_marker in config."
        )
    
    def _extract_pos_from_rows(self, stop_row: int) -> dict[str, int]:
        """Extract PO numbers from rows between header and stop marker."""
        pos = {}
        row = self.header_row + 1
        
        while row < stop_row:
            cell_value = self.sheet[f"{self.po_col}{row}"].value
            
            if self._is_valid_po(cell_value):
                po_number = str(cell_value).strip()
                pos[po_number] = row
            
            row += 1
        
        return pos
    
    def _is_valid_po(self, cell_value) -> bool:
        """Check if cell value represents a valid PO number."""
        if cell_value is None:
            return False
        
        po_str = str(cell_value).strip().lower()
        return po_str != "" and po_str != "none"
    
    def _log_pos_summary(self, pos: dict[str, int]) -> None:
        """Log summary of POs found in template."""
        if not pos:
            print("WARNING: No POs found in template.")
        else:
            print(f"Found {len(pos)} POs in template.")