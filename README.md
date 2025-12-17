# financial-automation-project
This repository contains an end-to-end ETL pipeline for automating financial data processing. This project automates the process of populating a Financial Spreadsheet Template with forecast data from vendor files and actual data from a C-TIES file.

## Key Inputs:

1. Vendor Forecast File: Contains monthly forecasted fees for each PO.

2. Transactional Detail File (C-TIES): Actuals, accruals, reversals by accounting period.

3. Financial Spreadsheet Template: Target file for writing data.

## Key Outputs:

1. Populated Financial Spreadsheet Template: Populated with data from the two input spreadsheets (forecasts, accruals, actuals) 

## Pipeline Flow:

1. Read vendor forecast file, transactional detail file

2. Parse data and store in JSON-like format:
```
{
    "PO12345": {
        "Jan": {
            "Accrual Reversal": 5000, 
            "Forecast": 2000
            "Accrual": 2500,
            "Actual": 3000
        }
        "Feb": {
            "Accrual Reversal": 5000, 
            "Forecast": 2000
            "Accrual": 2500,
            "Actual": 3000
        }
        ...
    }
}
```
3. Write to forecasting template and export


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
|
|── main.py                  # Orchestrates the pipeline
|
└── README.md
```

## Power Automate Integration

While this python script will handle all the complex data processing, we will levereage Power Automate to orchestrate the entire pipeline, resulting in nearly full automation. Power Automate will handle triggering, file movement, and notifications, and Python will handle complex logic and Excel manipulation (which is hard to maintain in Power Automate alone). 

**Full Automation Flow:**

1. Power Automate: Monitors SharePoint for new files, triggers the process, and orchestrates the workflow.

2. Python Script: Handles the heavy data processing—reading the forecast file, mapping months and POs, writing forecast values 
into the financial template, and saving the updated file.

3. SharePoint (Output Files): The processed template is uploaded back to SharePoint via Power Automate.

4. Notification: Power Automate sends a completion message (via email or Teams) with a link to the updated file.