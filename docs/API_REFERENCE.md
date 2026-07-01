# API Reference & Developer Guide

> **Complete API documentation and developer guide for the Financial Automation Project**

---

##  Table of Contents

- [Developer Setup](#developer-setup)
- [Project Structure](#project-structure)
- [Core API Reference](#core-api-reference)
- [Streamlit Components](#streamlit-components)
- [Utility Functions](#utility-functions)
- [Data Models](#data-models)
- [Extending the System](#extending-the-system)
- [Code Conventions](#code-conventions)
- [Testing Guidelines](#testing-guidelines)
- [Contributing](#contributing)

---

##  Developer Setup

### Prerequisites

- Python 3.9 or higher
- pip package manager
- Git
- Code editor (VS Code recommended)

### Environment Setup

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd financial-automation-project
   ```

2. **Create virtual environment** (recommended):
   ```bash
   python -m venv venv
   
   # Windows
   venv\Scripts\activate
   
   # Mac/Linux
   source venv/bin/activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Verify installation**:
   ```bash
   python -c "import pandas, openpyxl, streamlit; print('Setup complete!')"
   ```

### Development Tools

**Recommended VS Code Extensions**:
- Python
- Pylance
- YAML
- Excel Viewer

**Useful Commands**:
```bash
# Run Streamlit app
streamlit run app.py

# Run command-line pipeline
python main.py

# Run with specific config
python main.py --config configs/config_custom.yaml

# Format code (if using black)
black src/

# Type checking (if using mypy)
mypy src/
```

---

##  Project Structure

```
financial-automation-project/

 src/                           # Core source code
    __init__.py
    forecast_reader.py         # Forecast file reader
    transactional_detail_reader.py  # TIES reader
    template_reader.py         # Template structure reader
    template_writer.py         # Output workbook writer
    models.py                  # Data models
    utils.py                   # Utility functions

 app.py                         # Streamlit UI
 streamlit_backend.py           # Streamlit backend orchestration
 streamlit_config.py            # Streamlit configuration management
 main.py                        # Command-line entry point

 configs/                       # Configuration files
 data/                          # Sample data
 notebooks/                     # Jupyter notebooks
 docs/                          # Documentation
```

---

##  Core API Reference

### ForecastReader

**Location**: `src/forecast_reader.py`

**Purpose**: Read and process vendor forecast files

#### Constructor

```python
ForecastReader(file_paths: list, po_col: str)
```

**Parameters**:
- `file_paths` (list): List of paths to forecast Excel files
- `po_col` (str): Column name containing PO numbers (e.g., "PO #")

**Example**:
```python
from src.forecast_reader import ForecastReader

reader = ForecastReader(
    file_paths=[
        'data/forecasts/vendor1.xlsx',
        'data/forecasts/vendor2.xlsx'
    ],
    po_col='PO #'
)
```

---

#### Methods

##### load_forecast()

```python
def load_forecast(self) -> None
```

**Description**: Loads all forecast files into memory

**Behavior**:
- Finds valid sheets (must have PO column and forecast columns)
- Cleans PO numbers (converts floats to strings)
- Handles duplicate POs across files (first occurrence wins)
- Concatenates all data into single DataFrame

**Side Effects**:
- Sets `self.data` to combined DataFrame
- Prints warnings for duplicate POs

**Example**:
```python
reader.load_forecast()
# Prints: "Selected sheet 'Forecast' from file vendor1.xlsx"
# Prints: "WARNING: PO(s) PO12345 appear in multiple forecast files..."
```

---

##### get_forecast_data()

```python
def get_forecast_data(self) -> dict
```

**Description**: Extract and structure forecast data

**Returns**: Dictionary with structure:
```python
{
    'PO12345': {
        'Jan': {'Forecast': 1000.0, 'Source': [45, 46]},
        'Feb': {'Forecast': 2000.0, 'Source': [47]},
        # ... other months
    }
}
```

**Behavior**:
- Calls `load_forecast()` if data not already loaded
- Identifies forecast columns (ending with "- FTotal")
- Normalizes month names (first 3 letters)
- Groups by PO and sums values
- Tracks source rows for audit trail

**Example**:
```python
forecast_data = reader.get_forecast_data()

# Access specific PO and month
jan_forecast = forecast_data['PO12345']['Jan']['Forecast']  # 1000.0
source_rows = forecast_data['PO12345']['Jan']['Source']     # [45, 46]
```

---

##### _load_valid_sheet() (Internal)

```python
def _load_valid_sheet(self, file_path: str) -> Optional[pd.DataFrame]
```

**Description**: Finds and loads the first valid sheet from a file

**Parameters**:
- `file_path` (str): Path to Excel file

**Returns**: DataFrame if valid sheet found, None otherwise

**Valid Sheet Criteria**:
- Must contain `po_col` column
- Must contain at least one column ending with "- FTotal"

---

### TransactionalDetailReader

**Location**: `src/transactional_detail_reader.py`

**Purpose**: Read and categorize TIES transactional data

#### Constructor

```python
TransactionalDetailReader(
    file_path: str,
    required_cols: list,
    valid_types: list,
    colmap: dict
)
```

**Parameters**:
- `file_path` (str): Path to TIES Excel file
- `required_cols` (list): Columns required for valid sheet detection
- `valid_types` (list): Transaction types to process
- `colmap` (dict): Column name mappings

**Example**:
```python
from src.transactional_detail_reader import TransactionalDetailReader

reader = TransactionalDetailReader(
    file_path='data/transactional/cties.xlsx',
    required_cols=['PO Number', 'Month', 'GL Transaction Amount'],
    valid_types=['Actual', 'Accrual', 'Reversal'],
    colmap={
        'po': 'PO Number',
        'month': 'Month',
        'amount': 'GL BER Corp Amount',
        'classifier': 'AP Voucher Number',
        'cost_center': 'Cost Center*',
        'wbs': 'WBS Element',
        'type': 'Type'
    }
)
```

---

#### Methods

##### load_transactional_detail_file()

```python
def load_transactional_detail_file(self) -> None
```

**Description**: Loads all valid sheets from TIES file

**Behavior**:
- Checks each sheet for required columns
- Loads all valid sheets (header at row 2)
- Concatenates into single DataFrame
- Applies transaction categorization
- Converts PO numbers to strings

**Side Effects**:
- Sets `self.data` to combined DataFrame
- Adds 'Type' column with categorization
- Prints sheet names loaded

**Example**:
```python
reader.load_transactional_detail_file()
# Prints: "Loading valid sheets: ['AP01', 'AP02', 'AP03']"
# Prints: "Successfully loaded transactional data from valid sheets."
```

---

##### get_transactional_data()

```python
def get_transactional_data(self) -> dict
```

**Description**: Extract and aggregate transactional data

**Returns**: Dictionary with structure:
```python
{
    'PO12345': {
        'cost_center': '1234',
        'wbs': 'IT-CT123',
        'Jan': {
            'Actual': 900.0,
            'Accrual': 950.0,
            'Reversal': 0.0
        },
        # ... other months
    }
}
```

**Behavior**:
- Calls `load_transactional_detail_file()` if not loaded
- Filters to valid transaction types
- Groups by PO, Month, Type, Cost Center, WBS
- Applies month offset for actuals (actuals in AP02  Jan)
- Sorts months chronologically

**Month Offset Logic**:
- **Actuals**: Belong to prior month (AP02 actuals  Jan)
- **Accruals/Reversals**: Belong to current month (AP02 accruals  Feb)

**Example**:
```python
trans_data = reader.get_transactional_data()

# Access specific PO and month
jan_actual = trans_data['PO12345']['Jan']['Actual']      # 900.0
feb_accrual = trans_data['PO12345']['Feb']['Accrual']    # 950.0
cost_center = trans_data['PO12345']['cost_center']       # '1234'
```

---

##### get_hierarchy_map()

```python
def get_hierarchy_map(self) -> dict
```

**Description**: Returns row-level mapping for hierarchy building

**Returns**: Dictionary with structure:
```python
{
    45: {'po': 'PO12345', 'cost_center': '1234', 'wbs': 'IT-CT123'},
    46: {'po': None, 'cost_center': '2345', 'wbs': None},
    # ... all rows
}
```

**Behavior**:
- Maps every row in transactional file
- Normalizes missing values to None
- Validates row count matches DataFrame
- Prints summary statistics

**Example**:
```python
hierarchy_map = reader.get_hierarchy_map()

# Access specific row
row_45 = hierarchy_map[45]
po = row_45['po']              # 'PO12345'
wbs = row_45['wbs']            # 'IT-CT123'
cost_center = row_45['cost_center']  # '1234'
```

---

##### _categorize_row() (Internal)

```python
def _categorize_row(self, row: pd.Series) -> str
```

**Description**: Categorizes transaction type based on classifier and amount

**Parameters**:
- `row` (pd.Series): DataFrame row

**Returns**: Transaction type string

**Classification Rules**:
| Classifier Prefix | Amount | Type |
|-------------------|--------|------|
| 5xx | Any | "Actual" |
| 2xx | Positive | "Accrual" |
| 2xx | Negative | "Reversal" |
| 9xx | Any | "Reclass" |
| Other | Any | "Undefined" |

**Example**:
```python
# Applied automatically during load
# Row with AP Voucher "510123"  "Actual"
# Row with AP Voucher "210456" and amount 950  "Accrual"
# Row with AP Voucher "210789" and amount -950  "Reversal"
```

---

### TemplateReader

**Location**: `src/template_reader.py`

**Purpose**: Extract structure and PO mapping from template

#### Constructor

```python
TemplateReader(
    file_path: str,
    header_row: int,
    po_col: str,
    po_stop_marker: str,
    cost_center_col: str,
    cost_center_start_row: int
)
```

**Parameters**:
- `file_path` (str): Path to template Excel file
- `header_row` (int): Row number where PO headers start (1-based)
- `po_col` (str): Column letter containing PO numbers
- `po_stop_marker` (str): Text marker indicating end of PO section
- `cost_center_col` (str): Column letter containing cost center IDs
- `cost_center_start_row` (int): Row number where cost centers start (1-based)

**Example**:
```python
from src.template_reader import TemplateReader

reader = TemplateReader(
    file_path='data/templates/template.xlsx',
    header_row=16,
    po_col='B',
    po_stop_marker='Previous Period Invoices',
    cost_center_col='A',
    cost_center_start_row=9
)

# Properties automatically populated
cost_centers = reader.cost_centers  # ['1234', '2345', ...]
pos = reader.pos                    # {'PO12345': 17, 'PO67890': 18, ...}
```

---

#### Properties

##### cost_centers

```python
cost_centers: list[str]
```

**Description**: List of cost center IDs extracted from template

**Populated By**: `get_existing_cost_centers()` called in `__init__`

**Example**:
```python
reader = TemplateReader(...)
print(reader.cost_centers)  # ['1234', '2345', '3456']
```

---

##### pos

```python
pos: dict[str, int]
```

**Description**: Mapping of PO numbers to row numbers in template

**Populated By**: `get_existing_pos()` called in `__init__`

**Example**:
```python
reader = TemplateReader(...)
print(reader.pos)  # {'PO12345': 17, 'PO67890': 18, ...}

# Get row number for specific PO
row = reader.pos['PO12345']  # 17
```

---

#### Methods

##### get_existing_cost_centers()

```python
def get_existing_cost_centers(self) -> list[str]
```

**Description**: Reads cost centers from template

**Returns**: List of cost center IDs

**Behavior**:
- Reads from `cost_center_col` starting at `cost_center_start_row`
- Stops at first blank cell
- Strips whitespace
- Prints count found

**Example**:
```python
cost_centers = reader.get_existing_cost_centers()
# Prints: "Found 15 cost centers: ['1234', '2345', ...]"
```

---

##### get_existing_pos()

```python
def get_existing_pos(self) -> dict[str, int]
```

**Description**: Extracts PO numbers and their row positions

**Returns**: Dictionary mapping PO  row number

**Behavior**:
- Finds stop marker row
- Extracts POs between header row and stop marker
- Validates PO values (not None, not "none")
- Prints count found

**Raises**: `ValueError` if stop marker not found

**Example**:
```python
pos = reader.get_existing_pos()
# Prints: "Found 150 POs in template."
```

---

### TemplateWriter

**Location**: `src/template_writer.py`

**Purpose**: Generate output workbook with data and reports

#### Constructor

```python
TemplateWriter(
    file_path: str,
    output_path: str,
    overwrite: bool,
    header_row: int,
    po_column: str,
    dec_acc_reversal_col: str,
    forecast_source_cols: list,
    transactional_source_cols: list
)
```

**Parameters**:
- `file_path` (str): Input template path
- `output_path` (str): Output file path
- `overwrite` (bool): Whether to overwrite existing data
- `header_row` (int): Header row number (1-based)
- `po_column` (str): PO column letter
- `dec_acc_reversal_col` (str): Starting column for data
- `forecast_source_cols` (list): Columns for forecast audit sheet
- `transactional_source_cols` (list): Columns for transaction audit sheet

**Example**:
```python
from src.template_writer import TemplateWriter

writer = TemplateWriter(
    file_path='data/templates/template.xlsx',
    output_path='data/output/result.xlsx',
    overwrite=False,
    header_row=16,
    po_column='B',
    dec_acc_reversal_col='N',
    forecast_source_cols=['PO #', 'Jan 2026 - FTotal', ...],
    transactional_source_cols=['PO Number', 'Month', 'Type', ...]
)
```

---

#### Methods

##### write_hierarchy()

```python
def write_hierarchy(self, hierarchy: dict, pos: dict[str, int]) -> None
```

**Description**: Writes financial data to template

**Parameters**:
- `hierarchy` (dict): Cost Center  WBS  PO hierarchy
- `pos` (dict): PO  row number mapping

**Behavior**:
- Iterates through hierarchy (CostCenter  WBSCode  PO)
- For each PO, writes monthly metrics to correct cells
- Respects `overwrite` setting
- Skips POs not in template

**Example**:
```python
writer.write_hierarchy(hierarchy, pos=template_reader.pos)
```

---

##### write_forecast_source_sheet()

```python
def write_forecast_source_sheet(
    self,
    forecast_df: pd.DataFrame,
    pos: dict[str, int]
) -> None
```

**Description**: Creates forecast audit trail sheet

**Parameters**:
- `forecast_df` (pd.DataFrame): Forecast data
- `pos` (dict): PO  row mapping (for filtering)

**Behavior**:
- Filters to POs in template
- Creates new sheet "Forecast Source Data"
- Writes visible columns, hides others
- Adds total row with SUBTOTAL formulas
- Applies auto-filter and freeze panes
- Auto-sizes columns

**Example**:
```python
writer.write_forecast_source_sheet(
    forecast_reader.data,
    pos=template_reader.pos
)
```

---

##### write_transactional_source_sheet()

```python
def write_transactional_source_sheet(
    self,
    transactions_df: pd.DataFrame,
    pos: dict[str, int]
) -> None
```

**Description**: Creates transaction audit trail sheet

**Parameters**:
- `transactions_df` (pd.DataFrame): Transactional data
- `pos` (dict): PO  row mapping (for filtering)

**Behavior**:
- Filters to POs in template
- Creates new sheet "Transactions Source Data"
- Writes visible columns, hides others
- Applies auto-filter and freeze panes
- Auto-sizes columns

**Example**:
```python
writer.write_transactional_source_sheet(
    transactional_reader.data,
    pos=template_reader.pos
)
```

---

##### write_exception_sheet()

```python
def write_exception_sheet(
    self,
    exception_log: ExceptionLog,
    transactional_df: pd.DataFrame
) -> None
```

**Description**: Creates detailed exception log sheet

**Parameters**:
- `exception_log` (ExceptionLog): Exception log with entries
- `transactional_df` (pd.DataFrame): Source data for hidden columns

**Behavior**:
- Creates new sheet "Exceptions"
- Writes visible columns (Cost Center, Month, WBS, PO, Exception Type, Source Row, Amount, Type)
- Writes hidden columns (full source row data)
- Groups and hides supplementary columns
- Applies auto-filter and freeze panes

**Example**:
```python
writer.write_exception_sheet(
    exception_log,
    transactional_reader.data
)
```

---

##### write_exception_summary_sheet()

```python
def write_exception_summary_sheet(self, exception_log: ExceptionLog) -> None
```

**Description**: Creates executive exception summary with interactive filter

**Parameters**:
- `exception_log` (ExceptionLog): Exception log

**Behavior**:
- Creates new sheet "Exceptions Summary"
- Adds month filter dropdown
- Creates summary by exception type (with formulas)
- Creates summary by cost center (with formulas)
- Formulas update based on month filter selection

**Example**:
```python
writer.write_exception_summary_sheet(exception_log)
```

---

##### save()

```python
def save(self) -> None
```

**Description**: Saves workbook to output path

**Raises**: `Exception` if save fails

**Example**:
```python
writer.save()
# Prints: "Workbook saved to: data/output/result.xlsx"
```

---

##  Streamlit Components

### FileHandler

**Location**: `streamlit_backend.py`

**Purpose**: Handle temporary file storage for uploaded files

#### Methods

##### save_uploaded_files()

```python
@staticmethod
def save_uploaded_files(
    template_file,
    forecast_files: List,
    transactional_file
) -> Tuple[str, Dict[str, str]]
```

**Description**: Saves uploaded files to temporary directory

**Parameters**:
- `template_file`: Streamlit UploadedFile object
- `forecast_files` (List): List of Streamlit UploadedFile objects
- `transactional_file`: Streamlit UploadedFile object

**Returns**: Tuple of (temp_dir_path, file_paths_dict)

**Example**:
```python
temp_dir, file_paths = FileHandler.save_uploaded_files(
    template_file=template_file,
    forecast_files=forecast_files,
    transactional_file=transactional_file
)

# file_paths = {
#     'template': '/tmp/streamlit_finance_xxx/template.xlsx',
#     'forecast': ['/tmp/streamlit_finance_xxx/forecast_0_vendor1.xlsx', ...],
#     'transactional': '/tmp/streamlit_finance_xxx/cties.xlsx'
# }
```

---

##### cleanup_temp_files()

```python
@staticmethod
def cleanup_temp_files(temp_dir: str) -> None
```

**Description**: Removes temporary directory and contents

**Parameters**:
- `temp_dir` (str): Path to temporary directory

**Example**:
```python
FileHandler.cleanup_temp_files(temp_dir)
```

---

### PipelineOrchestrator

**Location**: `streamlit_backend.py`

**Purpose**: Orchestrates pipeline execution for Streamlit UI

#### Constructor

```python
PipelineOrchestrator(
    config: Dict,
    logger: Optional[StreamlitLogger] = None,
    progress_callback = None
)
```

**Parameters**:
- `config` (Dict): Configuration dictionary
- `logger` (Optional[StreamlitLogger]): Logger instance
- `progress_callback`: Function to call with progress updates (0-100)

**Example**:
```python
from streamlit_backend import PipelineOrchestrator, StreamlitLogger

logger = StreamlitLogger(status_container=st.empty())
orchestrator = PipelineOrchestrator(
    config=config_dict,
    logger=logger,
    progress_callback=lambda pct: progress_bar.progress(pct/100)
)
```

---

#### Methods

##### run()

```python
def run(
    self,
    file_paths: Dict[str, str],
    selected_cost_centers: Optional[List[str]] = None
) -> str
```

**Description**: Executes complete pipeline

**Parameters**:
- `file_paths` (Dict): Dictionary with 'template', 'forecast', 'transactional' keys
- `selected_cost_centers` (Optional[List[str]]): Cost centers to process

**Returns**: Path to generated output file

**Raises**: `Exception` if pipeline fails

**Example**:
```python
output_path = orchestrator.run(
    file_paths=file_paths,
    selected_cost_centers=['1234', '2345']
)
```

---

### ConfigManager

**Location**: `streamlit_config.py`

**Purpose**: Manages configuration loading, saving, and validation

#### Methods

##### load_config()

```python
@staticmethod
def load_config() -> AppConfig
```

**Description**: Loads configuration with priority

**Returns**: AppConfig object

**Priority**:
1. config_streamlit.yaml (if exists)
2. Default values

**Example**:
```python
from streamlit_config import ConfigManager

config = ConfigManager.load_config()
```

---

##### save_config()

```python
@staticmethod
def save_config(config: AppConfig) -> bool
```

**Description**: Saves configuration to config_streamlit.yaml

**Parameters**:
- `config` (AppConfig): Configuration to save

**Returns**: True if successful, False otherwise

**Example**:
```python
success = ConfigManager.save_config(config)
```

---

##### validate_config()

```python
@staticmethod
def validate_config(config: AppConfig) -> Tuple[bool, List[str]]
```

**Description**: Validates configuration values

**Parameters**:
- `config` (AppConfig): Configuration to validate

**Returns**: Tuple of (is_valid, list_of_errors)

**Example**:
```python
is_valid, errors = ConfigManager.validate_config(config)
if not is_valid:
    for error in errors:
        print(f"Error: {error}")
```

---

##  Utility Functions

### build_hierarchy()

**Location**: `src/utils.py`

```python
def build_hierarchy(
    cost_centers: list[str],
    hierarchy_map: dict,
    transactional_data: dict,
    forecast_data: dict,
    exception_log: ExceptionLog,
    transactional_df: pd.DataFrame
) -> dict[str, CostCenter]
```

**Description**: Builds Cost Center  WBS  PO hierarchy with exception tracking

**Parameters**:
- `cost_centers` (list[str]): List of cost center IDs to process
- `hierarchy_map` (dict): Row-level mapping from TransactionalDetailReader
- `transactional_data` (dict): Aggregated transaction data
- `forecast_data` (dict): Forecast data
- `exception_log` (ExceptionLog): Exception log for tracking issues
- `transactional_df` (pd.DataFrame): Source DataFrame for exception context

**Returns**: Dictionary mapping cost_center_id  CostCenter object

**Example**:
```python
from src.utils import build_hierarchy
from src.models import ExceptionLog

exception_log = ExceptionLog()

hierarchy = build_hierarchy(
    cost_centers=template_reader.cost_centers,
    hierarchy_map=transactional_reader.get_hierarchy_map(),
    transactional_data=transactional_reader.get_transactional_data(),
    forecast_data=forecast_reader.get_forecast_data(),
    exception_log=exception_log,
    transactional_df=transactional_reader.data
)

# Access hierarchy
cost_center = hierarchy['1234']
wbs = cost_center.wbs_codes['IT-CT123']
po = wbs.pos['PO12345']
jan_metrics = po.monthly_data['Jan']
```

---

### load_config()

**Location**: `src/utils.py`

```python
def load_config(config_path: str = 'configs/config_base.yaml') -> dict
```

**Description**: Loads YAML configuration file

**Parameters**:
- `config_path` (str): Path to YAML config file

**Returns**: Configuration dictionary

**Example**:
```python
from src.utils import load_config

config = load_config('configs/config_base.yaml')
```

---

### convert_base64()

**Location**: `src/utils.py`

```python
def convert_base64(bytes_string: str)
```

**Description**: Converts base64 string to Excel-like object

**Parameters**:
- `bytes_string` (str): Base64-encoded file content

**Returns**: BytesIO object readable by pandas/openpyxl

**Example**:
```python
from src.utils import convert_base64

# For Power Automate integration
excel_file = convert_base64(base64_string)
df = pd.read_excel(excel_file)
```

---

##  Data Models

Complete reference in [Architecture Guide - Data Models](ARCHITECTURE.md#data-models)

### Quick Reference

```python
from src.models import (
    MonthlyMetrics,
    PO,
    WBSCode,
    CostCenter,
    ExceptionType,
    ExceptionEntry,
    ExceptionLog
)

# Create hierarchy
metrics = MonthlyMetrics(forecast=1000, actual=900, accrual=950, accrual_reversal=0)
po = PO(po_number='PO12345', monthly_data={'Jan': metrics})
wbs = WBSCode(wbs_code='IT-CT123', cost_center='1234', pos={'PO12345': po})
cc = CostCenter(cost_center_id='1234', wbs_codes={'IT-CT123': wbs})

# Log exception
exception_log = ExceptionLog()
exception_log.log(
    ExceptionType.MISSING_WBS,
    row_index=42,
    po='PO12345',
    cost_center='1234',
    source_row_data={'PO Number': 'PO12345', ...}
)

# Get summaries
summary_by_type = exception_log.summary_by_type()
summary_by_cc = exception_log.summary_by_cost_center()
```

---

##  Extending the System

### Adding a New Exception Type

**Step 1**: Add to ExceptionType enum in `src/models.py`:
```python
class ExceptionType(Enum):
    MISSING_WBS_AND_PO = "MISSING_WBS_AND_PO"
    MISSING_WBS = "MISSING_WBS"
    MISSING_PO = "MISSING_PO"
    DUPLICATE_PO = "DUPLICATE_PO"
    DUPLICATE_WBS = "DUPLICATE_WBS"
    INVALID_AMOUNT = "INVALID_AMOUNT"  # New exception type
```

**Step 2**: Add detection logic in `src/utils.py` `build_hierarchy()`:
```python
# In build_hierarchy function
if amount < 0 and trans_type == 'Actual':
    exception_log.log(
        ExceptionType.INVALID_AMOUNT,
        row_index=row_idx,
        po=po,
        wbs=wbs,
        cost_center=cc_id,
        month=month,
        amount=amount,
        transaction_type=trans_type,
        source_row_data=source_row_data
    )
    continue
```

**Step 3**: Exception automatically appears in reports (no writer changes needed)

---

### Adding a New Data Source

**Step 1**: Create new reader class in `src/`:
```python
# src/budget_reader.py
import pandas as pd

class BudgetReader:
    def __init__(self, file_path: str, po_col: str):
        self.file_path = file_path
        self.po_col = po_col
        self.data = None
    
    def load_budget(self):
        self.data = pd.read_excel(self.file_path)
    
    def get_budget_data(self) -> dict:
        if self.data is None:
            self.load_budget()
        
        result = {}
        for _, row in self.data.iterrows():
            po = str(row[self.po_col])
            result[po] = {
                'Jan': {'Budget': row['Jan Budget']},
                'Feb': {'Budget': row['Feb Budget']},
                # ... other months
            }
        return result
```

**Step 2**: Update `MonthlyMetrics` in `src/models.py`:
```python
@dataclass
class MonthlyMetrics:
    forecast: float = 0.0
    actual: float = 0.0
    accrual: float = 0.0
    accrual_reversal: float = 0.0
    budget: float = 0.0  # New field
```

**Step 3**: Update `build_hierarchy()` in `src/utils.py`:
```python
def build_hierarchy(
    cost_centers: list[str],
    hierarchy_map: dict,
    transactional_data: dict,
    forecast_data: dict,
    budget_data: dict,  # New parameter
    exception_log: ExceptionLog,
    transactional_df: pd.DataFrame
) -> dict[str, CostCenter]:
    # ... existing code ...
    
    # Add budget data
    if po in budget_data:
        for month, values in budget_data[po].items():
            if month not in po_obj.monthly_data:
                po_obj.monthly_data[month] = MonthlyMetrics()
            po_obj.monthly_data[month].budget = values.get('Budget', 0.0)
```

**Step 4**: Update `main.py` and `streamlit_backend.py` to use new reader

---

### Adding a New Metric

**Step 1**: Update `MonthlyMetrics` in `src/models.py`:
```python
@dataclass
class MonthlyMetrics:
    forecast: float = 0.0
    actual: float = 0.0
    accrual: float = 0.0
    accrual_reversal: float = 0.0
    variance: float = 0.0  # New metric
```

**Step 2**: Calculate metric in `build_hierarchy()`:
```python
# After populating forecast and actual
if month in po_obj.monthly_data:
    metrics = po_obj.monthly_data[month]
    metrics.variance = metrics.actual - metrics.forecast
```

**Step 3**: Update `TemplateWriter.get_column_map()` in `src/template_writer.py`:
```python
def get_column_map(self, starting_col):
    months = ["Jan", "Feb", "Mar", ...]
    metrics = ["Accrual Reversal", "Forecast", "Accrual", "Actual", "Variance"]  # Add new metric
    
    # ... rest of method
```

**Step 4**: Update `write_hierarchy()` to write new metric:
```python
values = {
    'Accrual Reversal': metrics.accrual_reversal,
    'Forecast': metrics.forecast,
    'Accrual': metrics.accrual,
    'Actual': metrics.actual,
    'Variance': metrics.variance  # Add new metric
}
```

---

##  Code Conventions

### Type Hints

Always use type hints for function parameters and return values:

```python
def process_data(
    file_path: str,
    config: dict,
    validate: bool = True
) -> dict[str, Any]:
    """Process data from file."""
    pass
```

### Dataclasses

Use dataclasses for data structures:

```python
from dataclasses import dataclass, field
from typing import Optional

@dataclass
class Entity:
    required_field: str
    optional_field: Optional[str] = None
    collection: dict = field(default_factory=dict)
```

### Docstrings

Use docstrings for all public functions and classes:

```python
def calculate_total(items: list[float]) -> float:
    """
    Calculate total of all items.
    
    Args:
        items: List of numeric values
        
    Returns:
        Sum of all items
        
    Raises:
        ValueError: If items is empty
    """
    if not items:
        raise ValueError("Items list cannot be empty")
    return sum(items)
```

### Error Handling

Use specific exceptions and provide context:

```python
try:
    data = load_file(path)
except FileNotFoundError:
    raise FileNotFoundError(f"Could not find file: {path}")
except Exception as e:
    raise Exception(f"Error loading file {path}: {str(e)}")
```

### Naming Conventions

- **Classes**: PascalCase (`ForecastReader`, `ExceptionLog`)
- **Functions/Methods**: snake_case (`get_forecast_data`, `build_hierarchy`)
- **Constants**: UPPER_SNAKE_CASE (`DEFAULT_CONFIG_PATH`)
- **Private Methods**: Leading underscore (`_categorize_row`, `_load_valid_sheet`)

---

##  Testing Guidelines

### Unit Testing

Test individual components:

```python
import pytest
from src.forecast_reader import ForecastReader

def test_forecast_reader_loads_data():
    reader = ForecastReader(
        file_paths=['tests/data/test_forecast.xlsx'],
        po_col='PO #'
    )
    data = reader.get_forecast_data()
    
    assert 'PO12345' in data
    assert data['PO12345']['Jan']['Forecast'] == 1000.0

def test_forecast_reader_handles_duplicates():
    reader = ForecastReader(
        file_paths=[
            'tests/data/forecast1.xlsx',
            'tests/data/forecast2.xlsx'  # Contains duplicate PO
        ],
        po_col='PO #'
    )
    data = reader.get_forecast_data()
    
    # Should use first occurrence
    assert data['PO12345']['Jan']['Forecast'] == 1000.0  # From forecast1
```

### Integration Testing

Test component interactions:

```python
def test_full_pipeline():
    # Setup
    forecast_reader = ForecastReader(...)
    transactional_reader = TransactionalDetailReader(...)
    template_reader = TemplateReader(...)
    
    # Execute
    forecast_data = forecast_reader.get_forecast_data()
    transactional_data = transactional_reader.get_transactional_data()
    hierarchy_map = transactional_reader.get_hierarchy_map()
    
    exception_log = ExceptionLog()
    hierarchy = build_hierarchy(
        cost_centers=template_reader.cost_centers,
        hierarchy_map=hierarchy_map,
        transactional_data=transactional_data,
        forecast_data=forecast_data,
        exception_log=exception_log,
        transactional_df=transactional_reader.data
    )
    
    # Verify
    assert '1234' in hierarchy
    assert len(exception_log.entries) == 0  # No exceptions expected
```

### Test Data

Create minimal test files:

```python
# tests/data/test_forecast.xlsx
# Simple Excel file with:
# - PO # column
# - Jan 2026 - FTotal column
# - Few rows of data

# tests/data/test_cties.xlsx
# Simple Excel file with:
# - Required columns
# - Few rows of different transaction types
```

---

##  Contributing

### Development Workflow

1. **Create feature branch**:
   ```bash
   git checkout -b feature/new-feature
   ```

2. **Make changes**:
   - Follow code conventions
   - Add type hints
   - Write docstrings
   - Add tests

3. **Test changes**:
   ```bash
   python -m pytest tests/
   ```

4. **Commit changes**:
   ```bash
   git add .
   git commit -m "Add new feature: description"
   ```

5. **Push and create PR**:
   ```bash
   git push origin feature/new-feature
   ```

### Code Review Checklist

-  Code follows conventions
-  Type hints present
-  Docstrings added
-  Tests written and passing
-  Documentation updated
-  No breaking changes (or documented)

### Documentation Updates

When making changes:
- Update relevant API documentation
- Update user guide if user-facing
- Update configuration guide if config changes
- Add examples for new features

---

##  Related Documentation

- **[User Guide](USER_GUIDE.md)** - End user documentation
- **[Architecture Guide](ARCHITECTURE.md)** - System design
- **[Configuration Guide](CONFIGURATION.md)** - Configuration reference
- **[Deployment Guide](DEPLOYMENT.md)** - Operations and troubleshooting

---

**Last Updated**: June 2026  
**Version**: 1.0