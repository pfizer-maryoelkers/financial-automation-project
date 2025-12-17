
import pandas as pd

class TransactionalDetailReader:
    """Reads transactional detail file and extracts accruals, actuals, and reversals."""

    def __init__(self, file_path: str):
        """Initialize with the transactional detail file path."""
        self.file_path = file_path
        self.data = None

    def load_transactional_detail(self):
        """Load the transactional Excel file into memory."""
        # TODO: Implement reading the correct sheet and header row
        pass

    def get_transactional_data(self) -> dict:
        """
        Extract PO and monthly accruals, actuals, and reversals.
        Returns:
            dict: {
                'PO12345': {'Jan': {'Accrual': 950, 'Actual': 900, 'Reversal': 950}, ...}
            }
        """
        # TODO: Implement extraction and aggregation logic
