from openpyxl import load_workbook

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
        self.sheet = self.wb.active

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
        pos = {}
        row = self.header_row + 1
        # Find stop row using configurable marker
        stop_row = None
        for search_row in range(1, self.sheet.max_row + 1):
            if self.sheet[f"A{search_row}"].value == self.po_stop_marker:
                stop_row = search_row
                break
        if stop_row is None:
            raise ValueError(
                f"Could not find '{self.po_stop_marker}' marker in template. "
                "Please check the template format or update po_stop_marker in config."
            )
        while row < stop_row:
            cell = self.sheet[f"{self.po_col}{row}"].value
            # Skip None/blank cells instead of writing them
            if cell is None or str(cell).strip() == "" or str(cell).strip().lower() == "none":
                row += 1
                continue
            pos[str(cell).strip()] = row
            row += 1
        if not pos:
            print("WARNING: No POs found in template.")
        else:
            print(f"Found {len(pos)} POs in template.")
        return pos