# Financial Automation Project

An end-to-end ETL pipeline for automating financial data processing. This project automates the process of populating a Financial Spreadsheet Template with forecast data from vendor files and actual data from a C-TIES transactional detail file, with comprehensive exception tracking and data quality validation.

## Key Inputs

1. **Vendor Forecast File(s)**: Contains monthly forecasted fees for each PO
2. **Transactional Detail File (C-TIES)**: Actuals, accruals, and reversals by accounting period
3. **Financial Spreadsheet Template**: Target file with PO structure and formatting

## Key Outputs

1. **Populated Financial Spreadsheet Template**: Main sheet with forecast, actual, accrual, and reversal data
2. **Forecast Source Data Sheet**: Audit trail for forecast data
3. **Transactions Source Data Sheet**: Audit trail for transactional data
4. **Exceptions Sheet**: Detailed log of data quality issues with full source context
5. **Exceptions Summary Sheet**: High-level overview of exceptions by type and cost center

## Project Strucure:

```
financial-automation/
│
├── data/                        # Sample input files
│   ├── template.xlsx
│   ├── vendor_forecast.xlsx
│   └── transactional_detail.xlsx
│
├── src/
│   ├── forecast_reader.py       # Reads vendor forecast file
│   ├── template_writer.py       # Writes data into template
│   ├── transactional_detail_reader.py  # Reads transactional detail file
│   ├── utils.py       # contains a few miscellaneous helper functions
|
|── main.py                  # Orchestrates the pipeline
|
└── README.md
```


## Pipeline Flow:

1. Read vendor forecast file, transactional detail file. 

2. Parse data, and store in JSON-like format

3. Combine data into singular JSON

4. Write to forecasting template and export


Below is a flowchart to visualize the full flow of the pipeline:

![Flowchart](data/flowchart.png)

## Key Components

The pipeline follows an ETL (Extract, Transform, Load) architecture:

1. **Readers (Extract)**: Load data from source files
2. **Transformers (Transform)**: Build hierarchy and validate data quality
3. **Writers (Load)**: Generate output workbook with data and reports

## 1. Readers (Extract)

### ForecastReader

This class reads vendor forecast files and stores forecast data in a dictionary with the following structure:

```
dict: { 'PO12345': 
    {'Jan': 
        {'Forecast': 1000}, 
    'Feb': 
        {'Forecast': 2000}, 
    ...
    } 
}
```

**Parameters:**
- `file_paths`: List of file paths (supports multiple forecast files)
- `po_col`: Column name for PO numbers

**Returns:** Dictionary with PO → Month → Forecast structure

### TransactionalDetailReader

This class reads a singular transactional detail file and extracts accruals, actuals, and reversals.

Includes the following methods:
- load_transactional_detail_file(): Loads CTIES file into singular dataframe
- _categorize_row(): internal method that categorizes row type as actual, accrual, etc.
This method is where we define the logic for adding the 'Type' column - can be extended in the future
- get_transactional_data(): reads dataframe and filters data, returns dict of data we need

Returns a dictionary in the following format:
```
dict: {
    'PO12345': {
        'Jan': {
            'Actual': 900
            'Accrual': 950.0,
            'Accrual Reversal': 0.0,
        },
        'Feb': {
            'Actual': 800
            'Accrual': 1050.0,
            'Accrual Reversal': -950,
        }
        ...
    }
    ...
}
```

**Parameters:**
- `file_path`: Path to C-TIES file
- `required_cols`: Columns required for valid sheet detection
- `valid_types`: Supported transaction types (Actual, Accrual, Reversal)
- `colmap`: Column name mappings for flexibility

**Returns:** Dictionary with PO → Month → Transaction data

**Key Methods:**
- `get_transactional_data()`: Returns aggregated transaction data
- `get_hierarchy_map()`: Returns row-level mapping for hierarchy building

**Transactional Detail File Rules:**

In order to accurately parse data from the transactional detail file, we have defined a list of classification criteria that can concretely define a row as either an invoice (actual), and accrual, or an accrual reversal. We have defined the rules to be the following:

1. If colmap['classifier] has a prefix of 5, the entry is an actual

2. If colmap['classifier] has a prefix of 2 and the value is positive, the entry is an accrual 

3. If colmap['classifier] has a prefix of 2 and the value is negative, the entry is an accrual reversal

4. If colmap['classifier] has a prefix of 9, then entry is a reclass

5. Otherwise, the entry is undefined

### TemplateReader

Reads the template structure to extract cost centers and PO numbers.

**Parameters:**
- `file_path`: Path to template file
- `header_row`: Row where data entry begins
- `po_col`: Column containing PO numbers
- `cost_center_col`: Column containing cost center IDs

**Returns:** Cost center list and PO mapping

## 2. Transformers (Transform)

### Hierarchy Building (`build_hierarchy` in utils.py)

The core transformation step that:
1. Organizes data into Cost Center → WBS → PO hierarchy
2. Validates data quality and tracks exceptions
3. Combines forecast and transactional data

**Exception Detection:**

The system tracks five types of exceptions (in priority order):

1. **MISSING_WBS_AND_PO**: Both WBS code and PO number are missing
2. **MISSING_WBS**: WBS code is missing
3. **MISSING_PO**: PO number is missing
4. **DUPLICATE_WBS**: WBS code appears under multiple cost centers (all occurrences logged)
5. **DUPLICATE_PO**: PO appears under multiple WBS/cost center combinations

**Data Models:**

The system uses dataclasses for type safety:
- `CostCenter`: Contains WBS codes
- `WBSCode`: Contains POs and cost center reference
- `PO`: Contains monthly metrics (forecast, actual, accrual, reversal)
- `ExceptionEntry`: Captures exception details with full source context


## 3. Writers (Load)

### TemplateWriter

Writes data to the template and generates audit/exception sheets.

**Parameters:**
- `file_path`: Input template path
- `output_path`: Output file path
- `overwrite`: Whether to overwrite existing data
- `header_row`: Data entry start row
- `po_column`: PO column letter
- `dec_acc_reversal_col`: First data entry column (reference point)

**Key Methods:**

1. `write_hierarchy()`: Writes financial data to main template
2. `write_forecast_source_sheet()`: Creates forecast audit trail
3. `write_transactional_source_sheet()`: Creates transaction audit trail
4. `write_exception_sheet()`: Creates detailed exception log with:
   - Visible columns: Cost Center, WBS, PO, Exception Type, Source Row, Month, Amount, Type
   - Hidden columns: Full source row data (grouped and collapsible)
5. `write_exception_summary_sheet()`: Creates executive summary with:
   - Exception counts by type (with percentages)
   - Exception breakdown by cost center

**Output Sheets (in order):**
1. Main Template (populated with data)
2. Forecast Source Data
3. Transactions Source Data
4. Exceptions (detailed log)
5. Exceptions Summary (high-level overview)

## Configuration Management

The pipeline uses YAML configuration files for flexibility. This allows easy adaptation to different file formats and structures without code changes.

**Configuration Structure:**

```yaml
# Template configuration (shared between reader and writer)
template:
  file_path: "data/templates/Ram 2026_Owen Testing.xlsx"
  header_row: 16
  po_col: "B"
  po_stop_marker: "Previous Period Invoices"
  cost_center_col: "A"
  cost_center_start_row: 9

# Forecast reader configuration
forecast_reader:
  file_paths:
    - "data/forecasts/Other_Vendors_Forecasts.xlsx"
    - "data/forecasts/2026-Feb-IBM Forecast_AP02.xlsx"
  po_col: "PO #"

# Transactional detail reader configuration
transactional_detail_reader:
  file_path: "data/transactional/TIES AP03 2026.xlsx"
  required_cols:
    - "PO Number"
    - "Month"
    - "GL Transaction Amount"
  valid_types:
    - "Actual"
    - "Accrual"
    - "Reversal"
  colmap:
    po: "PO Number"
    month: "Month"
    amount: "GL BER Corp Amount"
    classifier: "AP Voucher Number"
    cost_center: "Cost Center*"
    wbs: "WBS Element"
    type: "Type"

# Template writer configuration
template_writer:
  output_path: "data/templates/template_AP03.xlsx"
  overwrite: False
  dec_acc_reversal_col: "N"
  forecast_source_cols:
    - "PO #"
    - "Jan 2026 - FTotal"
    - "Feb 2026 - FTotal"
    # ... (monthly columns)
  transactional_source_cols:
    - "PO Number"
    - "Accounting Period"
    - "AP Voucher Number"
    - "Vendor Name"
    - "WBS Element"
    # ... (additional audit columns)
```

**Key Configuration Sections:**
- `template`: Template file structure and layout
- `forecast_reader`: Forecast file paths and column mappings
- `transactional_detail_reader`: C-TIES file structure and validation rules
- `template_writer`: Output settings and source sheet columns

## Usage

### Running the Pipeline

```bash
# Install dependencies
pip install -r requirements.txt

# Run the pipeline
python main.py
```

### Output

The pipeline generates a workbook with:
1. **Main Template**: Populated with forecast, actual, accrual, and reversal data
2. **Forecast Source Data**: Audit trail for forecast values
3. **Transactions Source Data**: Audit trail for transactional values
4. **Exceptions**: Detailed exception log with full source context
5. **Exceptions Summary**: Executive overview of data quality issues

### Exception Analysis

**Start with the Summary:**
1. Open "Exceptions Summary" sheet
2. Review exception counts by type
3. Identify cost centers with issues

**Drill into Details:**
1. Open "Exceptions" sheet
2. Use filters to focus on specific issues
3. Expand hidden columns for full source data
4. Use "Source Row" to trace back to original file

## Power Automate Integration

The Python pipeline integrates with Power Automate for end-to-end automation:

**Automation Flow:**

1. **Trigger**: User submits Microsoft Form with file uploads
   - Forecast file(s)
   - Transactional detail file
   - Template file

2. **File Processing**: Power Automate converts files to base64 strings

3. **Pipeline Execution**: Files passed to Python script (hosted on Azure)

4. **Output Generation**: Python returns completed workbook as base64

5. **File Storage**: Power Automate uploads to OneDrive/SharePoint

6. **Notification**: User receives completion message with file link

**Benefits:**
- Fully automated workflow
- No manual file handling
- Audit trail maintained
- Exception reports for data quality monitoring

## Recent Enhancements

### Exception Tracking System (May 2026)

Major improvements to exception detection and reporting:

**New Exception Types:**
- `MISSING_WBS_AND_PO`: Detects rows missing both identifiers
- `DUPLICATE_WBS`: Flags WBS codes owned by multiple cost centers

**Enhanced Exception Data:**
- Full source row context for every exception
- Month, amount, and transaction type captured
- All transactional columns available (grouped/hidden)

**New Reporting:**
- Exceptions Summary sheet with executive overview
- Exception counts by type and cost center
- Improved column naming ("Source Row" vs "Row Index")

**Removed:**
- `MISSING_FORECAST` exception (no longer tracked)

See `IMPLEMENTATION_SUMMARY.md` and `QUICK_REFERENCE.md` for detailed documentation.

## Documentation

- **README.md** (this file): Project overview and architecture
- **EXCEPTIONS_ENHANCEMENT_PLAN.md**: Detailed implementation plan for exception system
- **IMPLEMENTATION_SUMMARY.md**: Complete change summary and testing guide
- **QUICK_REFERENCE.md**: Quick guide for using exception reports

## Contributing

When making changes:
1. Update relevant configuration in `configs/`
2. Follow existing code patterns and type hints
3. Update documentation as needed
4. Test with sample data before deployment

## License

Internal Pfizer project - All rights reserved
