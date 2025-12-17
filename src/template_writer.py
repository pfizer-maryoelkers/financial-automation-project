
from openpyxl import load_workbook

class TemplateWriter:
    """Handles parsing template headers and writing forecast data."""

    def __init__(self, template_path: str):
        """Initialize with the template file path."""
        self.template_path = template_path
        self.workbook = None
        self.header_map = {}

    def parse_headers(self):
        """Parse template headers to map months and fields to columns."""
        # TODO: Implement header parsing logic
        pass

    def write_forecast(self, model: dict):
        """
        Write forecast data into the template.
        Args:
            model (dict): Data model with PO and monthly forecast values.
        """
        # TODO: Implement writing logic
        pass

    # TODO: Add support for transactional detail data 
    # Separate methods for acrruals, actuals, reversals?

    def save(self, output_path: str):
        """Save the updated template to the given path."""
        # TODO: Implement save logic