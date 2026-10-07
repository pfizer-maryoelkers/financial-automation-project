from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


class TemplateReader:
    """OpEx template reader.

    Reads the front-tab of a TIES / OpEx Excel template.  The front tab is
    the workbook's *active* sheet — the main data sheet that contains:

    - A cost-center list in a header section (col A, above the data rows)
    - A PO data table between the header row and a "Previous Period Invoices"
      stop-marker row

    This reader is **OpEx-specific**.  For project / CapEx templates (those
    with a P3 ID config sheet or a WBS/P3-ID column pair) use
    ``ProjectTemplateReader`` from ``src.project_template_reader`` instead.

    The pipeline separation is:
        OpEx  →  TemplateReader  +  build_hierarchy  +  TemplateWriter (no p3_id_column)
        Project → ProjectTemplateReader + build_project_hierarchy + TemplateWriter (p3_id_column set)
    """

    def __init__(self,
                file_path,
                header_row,
                po_col,
                po_stop_marker,
                cost_center_col,
                cost_center_start_row,
                cost_center_end_row=None,
                **kwargs
        ):

        self.wb = load_workbook(file_path)

        # OpEx templates store all data on their active (first) sheet.
        self.sheet: Worksheet = self.wb.active  # type: ignore[assignment]

        if self.sheet is None:
            raise ValueError(f"Could not load active sheet from {file_path}")

        sheet_name = self.sheet.title
        print(f"OpEx TemplateReader: opened sheet '{sheet_name}' from '{file_path}'")

        # Initializing instance variables
        self.header_row = header_row
        self.po_col = po_col
        self.po_stop_marker = po_stop_marker
        self.cost_center_col = cost_center_col
        self.cost_center_start_row = cost_center_start_row
        self.cost_center_end_row = cost_center_end_row

        # Read on init
        self.cost_centers = self.get_existing_cost_centers()
        self.template_rows = self.get_template_rows()
        self.pos = self.get_existing_pos()

    # ------------------------------------------------------------------
    # Cost center discovery
    # ------------------------------------------------------------------

    def _find_cost_center_start_row(self) -> int:
        """Scan cost_center_col for a 'Cost Center' header cell.

        Returns the row immediately after the header cell so iteration starts
        on the first actual cost-center value.  Falls back to the configured
        ``cost_center_start_row`` when the header is not found.
        """
        max_row = self.sheet.max_row or 1000
        for search_row in range(1, max_row + 1):
            cell_val = self.sheet[f"{self.cost_center_col}{search_row}"].value
            if cell_val is not None and "cost center" in str(cell_val).strip().lower():
                return search_row + 1
        print(
            f"WARNING: 'Cost Center' marker not found in column {self.cost_center_col}. "
            f"Falling back to configured cost_center_start_row={self.cost_center_start_row}."
        )
        return self.cost_center_start_row

    def get_existing_cost_centers(self) -> list[str]:
        """Return the list of cost centers declared in the template header section.

        Reads from ``cost_center_col`` starting immediately after the
        'Chargeout Cost Center' (or equivalent) header cell.  Stops when a
        blank cell or a cell whose text starts with 'Expense' is encountered.

        Returns
        -------
        list[str]
            e.g. ``['1234', '2345', 'CC-999']``
        """
        start_row = self._find_cost_center_start_row()
        cost_centers = []
        row = start_row

        while True:
            if self.cost_center_end_row is not None and row > self.cost_center_end_row:
                break
            cell = self.sheet[f"{self.cost_center_col}{row}"].value
            if cell is None or str(cell).strip() == "":
                break
            cell_text = str(cell).strip()
            if cell_text.lower().startswith("expense"):
                break
            cost_center = cell_text.split("/")[0].strip()
            cost_centers.append(cost_center)
            row += 1

        if not cost_centers:
            print("WARNING: No cost centers found in template.")
        else:
            print(f"OpEx TemplateReader: found {len(cost_centers)} cost center(s): {cost_centers}")
        return cost_centers

    # ------------------------------------------------------------------
    # PO / row position reading
    # ------------------------------------------------------------------

    def get_existing_pos(self) -> dict[str, int]:
        """Return a mapping of PO number → row index from the template data table."""
        pos = {po: data['row'] for po, data in self.template_rows.items()}
        self._log_pos_summary(pos)
        return pos

    def get_template_rows(self) -> dict[str, dict[str, str | int | None]]:
        """Return full PO metadata (row, cost_center) for every PO in the data table."""
        stop_row = self._find_stop_row()
        return self._extract_pos_from_rows(stop_row)

    def _find_stop_row(self) -> int:
        """Find the row containing the ``po_stop_marker`` text in column A.

        Returns ``max_row + 1`` (a safe sentinel) when the marker is absent so
        that ``_extract_pos_from_rows`` scans nothing instead of raising.
        """
        max_row = self.sheet.max_row or 1000
        for search_row in range(1, max_row + 1):
            if self.sheet[f"A{search_row}"].value == self.po_stop_marker:
                return search_row

        print(
            f"WARNING: '{self.po_stop_marker}' marker not found in sheet "
            f"'{self.sheet.title}' — treating as blank template with no existing POs."
        )
        return max_row + 1

    def _extract_pos_from_rows(self, stop_row: int) -> dict[str, dict[str, str | int | None]]:
        """Scan rows from ``header_row + 1`` up to (but not including) ``stop_row``.

        Tracks the running cost-center label from ``cost_center_col`` so each
        PO row is associated with the correct cost center.
        """
        pos = {}
        row = self.header_row + 1
        current_cost_center = None

        while row < stop_row:
            cc_value = self.sheet[f"{self.cost_center_col}{row}"].value
            if cc_value is not None:
                cc_text = str(cc_value).strip()
                if cc_text and not cc_text.lower().startswith("expense"):
                    current_cost_center = cc_text.split("/")[0].strip()

            cell_value = self.sheet[f"{self.po_col}{row}"].value

            if self._is_valid_po(cell_value):
                s = str(cell_value).strip()
                # Normalise float-formatted integers (e.g. 9500905777.0 → "9500905777")
                if s.replace('.', '', 1).replace('-', '', 1).isdigit():
                    try:
                        s = str(int(float(s)))
                    except (ValueError, OverflowError):
                        pass
                pos[s] = {
                    'row': row,
                    'cost_center': current_cost_center,
                }

            row += 1

        return pos

    def _is_valid_po(self, cell_value) -> bool:
        """Return True when ``cell_value`` looks like a real PO number."""
        if cell_value is None:
            return False
        po_str = str(cell_value).strip().lower()
        return po_str != "" and po_str != "none"

    def _log_pos_summary(self, pos: dict) -> None:
        if not pos:
            print("WARNING: No POs found in OpEx template.")
        else:
            print(f"OpEx TemplateReader: found {len(pos)} PO(s) in template.")
