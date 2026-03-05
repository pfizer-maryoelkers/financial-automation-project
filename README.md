# financial-automation-project
This repository contains an end-to-end ETL pipeline for automating financial data processing. This project automates the process of populating a Financial Spreadsheet Template with forecast data from vendor files and actual data from a C-TIES file.

## Key Inputs:

1. Vendor Forecast File: Contains monthly forecasted fees for each PO.

2. Transactional Detail File (C-TIES): Actuals, accruals, reversals by accounting period.

3. Financial Spreadsheet Template: Target file for writing data.

## Key Outputs:

1. Populated Financial Spreadsheet Template: Populated with data from the two input spreadsheets (forecasts, accruals, actuals) 

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

## Key Components:

The components can be grouped into three categories:

1. Readers (Extract)
2. Aggregators (Transform)
3. Writers (Load)

## 1. Readers (Extract):


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

Parameters: 
- file_paths -> list of file paths or byte strings (supports multiple)

Returns:
- dict

### Transactional Detail Reader

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

Parameters:
- file_path -> path or byte string
- required_cols -> list that defines the columns required for a valid sheet. Some CTIES files have multiple sheets, so this parameter helps dynamically identify which sheets to use
- valid_types -> list that defined supported types of transactions. For example Actual, Accrual, Reversal, etc. are supported types of transactions.
- colmap -> dict that normalizes column names. For example, dict['po'] may equal 'PO Number', 'PO #', etc. depending on config

Returns;
- dict

**Transactional Detail File Rules:**

In order to accurately parse data from the transactional detail file, we have defined a list of classification criteria that can concretely define a row as either an invoice (actual), and accrual, or an accrual reversal. We have defined the rules to be the following:

1. If colmap['classifier] has a prefix of 5, the entry is an actual

2. If colmap['classifier] has a prefix of 2 and the value is positive, the entry is an accrual 

3. If colmap['classifier] has a prefix of 2 and the value is negative, the entry is an accrual reversal

4. If colmap['classifier] has a prefix of 9, then entry is a reclass

5. Otherwise, the entry is undefined

## 2. Aggregators (Transform):

When we initially read the data, we have 2 separate dictionaries:
1. Forecast dict
2. Transactional Detail dict

This data needs to be aggregated to a singular JSON formatted dictionary so that it can be easily passed to our template writer. We do so with a helper function in utils.py:

```python
#  Function to combine forecasts, actuals, and accrual data into one JSON formatted dictionary
def combine_data(forecast, transactional):
    return combined_data
```

This function outputs a dictionary like the following:

```
dict: { 'PO12345': 
    {'Jan': 
        'Forecast': 1000,
        'Actual': 900,
        'Accrual': 950.0,
        'Accrual Reversal': 0.0,
    'Feb': 
        'Forecast': 1000,
        'Actual': 900,
        'Accrual': 950.0,
        'Accrual Reversal': -950.0,
    ...
    } 
}
```
This aggregated dictionary includes forecasts, actuals, accruals, and reversals in one place -- making it easier to write to our output file.


## 3. Writers (Load):

### Template Writer

This class reads an input template, parses the existing data, and appends the new data to the template. This class leverages both openpyxl and pandas for data manipulation/writing. 

Parameters:
- file_path -> file path to input template with POs and formatting
- header_row -> this is the row where data entry starts, generally row 14 but can be configured
- po_column -> this is the column where POs are entered, generally column B but can be configured
- dec_acc_reversal_col -> column where December Accrual Reversal exists. This is the very first data entry cell, so we use it as a reference point to build a map for the rest of the data writing

Key methods:
- write_data -> this method writes data to the main sheet
- write_forecast_source_sheet -> this method writes forecast data to a new sheet so that the main sheet can be easily audited. It is a trimmed down version of the actual forecast file
- write_transactional_source_sheet -> similarly, we write the transactional data to a seperate source sheet so that values can be easily audited.

Returns:
- Updated template excel file w/ filled in data and source sheets. Returns as either excel file or as byte string.

--------
## YAML Configuration

For increased flexibility, we have added YAML support for configurations for each class. This way, the user can easily update configs depending on the format of the forecast files, transactional detail files, and template files. Below is an example YAML config file:

```yaml
## Forecast Reader
forecast_reader:
  file_paths:
    - "data/2025-11-IBM Forecast.xlsx"

## Transactional Detail Reader
transactional_detail_reader:
  file_path: "data/C-TIES AP09 2025.xlsx"
  
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
    amount: "GL Transaction Amount"
    classifier: "AP Voucher Number"
    type: "Type"

## Template Writer
template_writer:
  file_path: "data/templates/financial_template_v1_2025.xlsx"
  header_row: 14
  po_column: "B"
  dec_acc_reversal_col: "J"

```
This YAML template includes all of the parameters that can possibly change depending on the inputs. Updating one of these parameters here ensures that the pipeline will run smoothly with varying configurations.

--------


## Power Automate Integration

While this python script will handle all the complex data processing, we will levereage Power Automate to orchestrate the entire pipeline, resulting in nearly full automation. Power Automate will handle triggering, file movement, and notifications, and Python will handle complex logic and Excel manipulation (which is hard to maintain in Power Automate alone). 

**Full Automation Flow:**

1. User submits Microsoft Form
    - This form includes 3 file uploads - Forecast File(s), Transactional Detail File, Template

2. Power Automate gathers files from OneDrive after form is submitted

3. Power Automate converts files to base63 encoded strings

4. The base64 excel files are passed over HTTPS as a JSON object which is taken in as input to main.py (hosted on Azure)

5. main.py is called and outputs the completed template as a base64 encoded string

6. This string is then passed back to Power Automate and uploaded to OneDrive/Sharepoint

7. Power Automate sends a completion message (via email or Teams) with a link to the updated file.
