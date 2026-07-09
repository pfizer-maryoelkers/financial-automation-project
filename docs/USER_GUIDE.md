# Financial Automation Portal — User Guide

> **Using Financial Template, TIES, and Vendor Forecast Files**

---

## Table of Contents

- [User Guide for Non-Technical Users](#user-guide-for-non-technical-users)
  - [Accessing the Portal](#accessing-the-portal)
  - [Before You Begin](#before-you-begin)
  - [Generating a Report](#generating-a-report)
  - [Understanding the Output Workbook](#understanding-the-output-workbook)
  - [Common Exception Types](#common-exception-types)
  - [Best Practices](#best-practices)
  - [Troubleshooting](#troubleshooting)
- [User Guide for Technical Users](#user-guide-for-technical-users)
  - [Getting Started](#getting-started)
  - [Input File Requirements](#input-file-requirements)
  - [Configuration Settings](#configuration-settings)
  - [Report Generation Process](#report-generation-process)
  - [Understanding Outputs](#understanding-outputs)
  - [Exception Management](#exception-management)
  - [Best Practices (Technical)](#best-practices-technical)
  - [Common Workflows](#common-workflows)
  - [Troubleshooting (Technical)](#troubleshooting-technical)

---

---

# User Guide for Non-Technical Users

> A step-by-step guide for generating financial reports using the Financial Automation Portal

---

## Accessing the Portal

Access the Financial Automation Portal using the following link:

👉 [Financial Automation Generator](https://financial-automation-generator.streamlit.app)

The portal allows users to upload source files, generate reports, review exceptions, and download the completed workbook.

---

## Before You Begin

Before generating a report, ensure you have the following files:

### Required Files

| File | Format |
|------|--------|
| Financial Template Workbook | `.xlsx` |
| TIES Transactional Detail File | `.xlsx` |
| One or more Forecast Files | `.xlsx` |

### Important Validation Checks

**Template Workbook**
- Confirm the worksheet structure is correct.
- Verify Cost Centers and PO information are present.
- Ensure column locations align with the portal configuration settings.

**TIES Transactional Detail File**
- Ensure all required columns are included.
- Month must be populated for every transaction.
- Verify Cost Center, WBS, PO Number, and transaction amount data are present.
- Confirm the file structure matches the configured column mappings.

**Forecast Files**
- Ensure forecast months are included.
- Verify PO numbers correspond to those in the template.
- Confirm file formatting aligns with the configuration settings.

> **Note:** Incorrect configurations or missing Month values (in TIES sheet) may result in missing data, inaccurate reporting, or processing errors.

---

## Generating a Report

### Step 1: Upload Files

Upload the required files in the following order:

1. Template Workbook
2. TIES Transactional Detail File
3. Forecast File(s)

After uploading, the portal validates the files and displays confirmation messages.

---

### Step 2: Review Cost Centers

When the template is uploaded, the system identifies available Cost Centers.

Users can:
- Process all Cost Centers
- Select specific Cost Centers
- Exclude Cost Centers not required for the current run

This option is useful when testing or reviewing a specific area.

---

### Step 3: Review Configuration Settings

Before generating the report, review the configuration section to ensure it matches the layout of the uploaded spreadsheets.

Key settings include:

**Template Configuration**
- Header Row
- Cost Center Column
- PO Column

**TIES Configuration**
- Month Column
- PO Number Column
- Transaction Amount Column
- Cost Center Column
- WBS Column

**Forecast Configuration**
- PO Number Column
- Monthly Forecast Columns

> **Best Practice:** Always review configuration settings when receiving updated file formats from vendors or business partners.

---

### Step 4: Generate the Report

Once all files have been uploaded and validated:

1. Confirm all uploads show a successful status.
2. Verify configuration settings.
3. Select **Generate Report**.

The system will process the files and display progress updates throughout the report generation process.

---

### Step 5: Review Exception Summary

After processing completes, an Exception Summary is displayed.

The summary provides:
- Total exception count
- Exception counts by category
- Percentage breakdown by exception type

| Exception Volume | Recommendation |
|-----------------|----------------|
| 0 Exceptions | Data quality is excellent |
| Few Exceptions | Review and determine if corrective action is required |
| High Exception Volume | Investigate source data and validate configurations |

---

### Step 6: Preview Results

Before downloading the workbook, preview the generated report.

Review:
- Cost Center information
- WBS assignments
- Purchase Orders
- Forecast amounts
- Actuals
- Accruals
- Reversals

Verify that the data appears as expected before downloading.

---

### Step 7: Download the Workbook

Once validation is complete:

1. Select **Download Report**.
2. Save the workbook locally.
3. Open in Excel for detailed analysis and distribution.

---

## Understanding the Output Workbook

The downloaded workbook contains several worksheets designed for reporting and auditing purposes.

### Main Template

Contains the completed financial report populated with:
- Forecast values
- Actuals
- Accruals
- Reversals

This is the primary reporting worksheet used for financial analysis.

### Forecast Source Data

Provides an audit trail of all forecast data used in the report.

Use this worksheet to:
- Verify forecast values
- Review vendor submissions
- Trace forecast data back to source files

### TIES Source Data

Provides a complete audit trail of transaction data loaded from TIES.

Use this worksheet to:
- Verify actuals
- Validate accruals and reversals
- Investigate discrepancies

### Exceptions Detail

Contains all exception records identified during processing.

Available information includes:
- Cost Center
- Month
- WBS
- PO Number
- Exception Type
- Source Row
- Transaction Amount

Use this worksheet to investigate and resolve data quality issues.

### Exceptions Summary

Provides a high-level overview of all exceptions, including:
- Exception counts by type
- Exception counts by Cost Center
- Percentage breakdowns
- Monthly filtering capabilities

This worksheet is useful for management reporting and data quality monitoring.

---

## Common Exception Types

### Missing WBS and PO
Both WBS and PO are missing from a transaction.
**Impact:** Transaction cannot be assigned and will not be included in reporting.

### Missing WBS
A PO exists, but the WBS value is missing.
**Impact:** Ownership of the transaction cannot be determined.

### Missing PO
A WBS exists, but the PO Number is missing.
**Impact:** Transaction cannot be mapped correctly.

### Duplicate WBS
The same WBS appears under multiple Cost Centers.
**Impact:** Ownership conflicts may occur.

### Duplicate PO
The same PO appears under multiple WBS or Cost Center combinations.
**Impact:** Duplicate records may be excluded from reporting.

---

## Best Practices

Before every report run:

- ✅ Verify the Month field is populated in the TIES Transactional Detail File
- ✅ Confirm all uploaded files are current versions
- ✅ Review configuration settings before processing
- ✅ Verify Cost Centers are selected appropriately
- ✅ Review exception reports after each run
- ✅ Investigate and resolve recurring data quality issues

---

## Troubleshooting

### Report Does Not Generate

- Verify all required files were uploaded.
- Confirm configuration settings match the spreadsheet layouts.
- Ensure the Month column exists within the TIES file.
- Check for missing required columns.

### Data Appears Missing

- Review the Exceptions worksheet.
- Verify PO and WBS values exist in the source files.
- Confirm Cost Centers are selected for processing.
- Recheck configuration mappings.

### High Number of Exceptions

- Review source data quality.
- Validate TIES data completeness.
- Confirm PO, WBS, Cost Center, and Month values are populated.
- Verify configuration settings match the uploaded files.

---

---

# User Guide for Technical Users

> Financial Automation of Template, TIES, and Vendor Files — Technical User Guide

**Application URL:** 👉 [Financial Automation Generator](https://financial-automation-generator.streamlit.app)

**GitHub:** 👉 [Financial Automation GitHub](https://github.com/your-org/financial-automation-project)

---

## Getting Started

### Prerequisites

Before using the Financial Automation solution, ensure you have:

- Access to the Financial Automation Portal
- Financial Template Workbook (`.xlsx`)
- TIES Transactional Detail File (`.xlsx`)
- One or more Vendor Forecast Files (`.xlsx`)

### Required Validation

Prior to processing files:

- Verify that the Month column exists in the TIES file.
- Ensure all transactions contain valid Month values.
- Confirm PO Numbers are populated where applicable.
- Verify Cost Centers and WBS Elements are present.
- Ensure file layouts match the expected configuration settings.

---

## Input File Requirements

### Template Workbook

The Template Workbook serves as the report structure that will be populated during processing.

**Requirements**
- Excel (`.xlsx`) format
- Contains Cost Center information
- Contains Purchase Order (PO) information
- Monthly Forecast, Accrual, Reversal, and Actual columns
- Consistent worksheet structure

**Validation Checklist**
- ✅ Cost Centers are current
- ✅ PO Numbers are populated
- ✅ Monthly columns exist
- ✅ Workbook structure has not been modified

### Vendor Forecast Files

Vendor Forecast files provide forecast amounts for Purchase Orders.

**Requirements**
- Excel (`.xlsx`) format
- Contains PO Number column
- Contains monthly forecast values
- Current reporting period data

**Validation Checklist**
- ✅ Forecast months are included
- ✅ PO Numbers match the template
- ✅ Duplicate PO values are reviewed
- ✅ Latest forecast version is being used

### TIES Transactional Detail File

The TIES file provides transaction data used to populate Actuals, Accruals, and Reversals.

**Required Fields**
- PO Number
- Month
- Cost Center
- WBS Element
- Transaction Amount
- AP Voucher Number

> **Critical Requirement:** The Month field must be included and populated. Missing Month values may prevent transactions from being processed correctly.

**Validation Checklist**
- ✅ Month column exists
- ✅ Month values are populated
- ✅ Cost Centers are populated
- ✅ WBS Elements are populated where applicable
- ✅ Transaction amounts are valid
- ✅ Latest TIES extract is being used

---

## Configuration Settings

Before generating a report, review configuration settings and confirm they match the uploaded files.

### Template Configuration

Verify:
- Header Row
- Cost Center Column
- PO Column
- Stop Marker

### TIES Configuration

Verify:
- Month Column
- PO Number Column
- Cost Center Column
- WBS Column
- Transaction Amount Column

### Forecast Configuration

Verify:
- PO Number Column
- Forecast Value Columns
- Forecast Worksheet Selection

> **Important:** Configuration mismatches are one of the most common causes of missing report data, incorrect report population, high exception volumes, and processing failures.

---

## Report Generation Process

### Step 1: Upload Files

Upload the following files:

1. Template Workbook
2. TIES Transactional Detail File
3. Vendor Forecast File(s)

The system validates uploaded files before processing.

### Step 2: Review Cost Centers

The system automatically extracts Cost Centers from the template.

Users may:
- Process all Cost Centers
- Select specific Cost Centers
- Temporarily exclude Cost Centers

### Step 3: Verify Configuration Settings

Review all configuration settings prior to running the report. Ensure uploaded file structures align with configured settings.

### Step 4: Generate Report

Select **Generate Report**. The system performs the following actions:

1. Loads source files
2. Processes forecast data
3. Processes transactional data
4. Builds Cost Center, WBS, and PO relationships
5. Populates the template
6. Generates exception reports
7. Creates the final workbook

### Step 5: Review Exception Summary

Upon completion, review the Exception Summary:

- Total exception count
- Exception categories
- Cost Centers with the most exceptions
- Trends requiring investigation

### Step 6: Download Report

Download the completed workbook for validation and distribution.

---

## Understanding Outputs

The generated workbook contains multiple worksheets.

### Main Template

Contains populated reporting data including Forecasts, Actuals, Accruals, and Reversals.

**Primary Use:** Financial reporting and analysis.

### Forecast Source Data

Contains forecast information used during report generation.

**Primary Use:** Forecast validation, audit support, source verification.

### TIES Source Data

Contains transactional information loaded from the TIES file.

**Primary Use:** Transaction validation, audit support, variance investigation.

### Exceptions Detail

Contains detailed exception records.

| Field | Description |
|-------|-------------|
| Cost Center | Cost center associated with the record |
| Month | Transaction month |
| WBS | WBS element |
| PO Number | Purchase order number |
| Exception Type | Category of exception |
| Source Row | Row number from source file |
| Transaction Amount | Transaction value |

**Primary Use:** Root cause investigation and issue resolution.

### Exceptions Summary

Provides a high-level overview of exception activity including exception counts, exception percentages, Cost Center breakdowns, and month-based filtering.

**Primary Use:** Data quality monitoring and management reporting.

---

## Exception Management

Exception reports identify records that could not be processed correctly.

### Missing WBS and PO

**Description:** Both the WBS and PO Number are missing.
**Impact:** Transaction cannot be assigned and is excluded from reporting.

### Missing WBS

**Description:** PO exists but WBS is missing.
**Impact:** Transaction ownership cannot be determined.

### Missing PO

**Description:** WBS exists but PO Number is missing.
**Impact:** Transaction cannot be assigned to a Purchase Order.

### Duplicate WBS

**Description:** The same WBS appears under multiple Cost Centers.
**Impact:** Ownership conflicts may occur.

### Duplicate PO

**Description:** The same PO appears under multiple Cost Center and WBS combinations.
**Impact:** Duplicate records may be excluded during processing.

### Exception Investigation Process

1. Open the **Exceptions Summary** worksheet.
2. Identify the largest exception categories.
3. Open the **Exceptions Detail** worksheet.
4. Filter by Exception Type.
5. Locate affected transactions.
6. Review source files.
7. Correct source data.
8. Re-run the report.
9. Verify exception reduction.

---

## Best Practices (Technical)

### File Preparation

Before uploading files:

- ✅ Verify all files are in Excel format (`.xlsx`)
- ✅ Ensure files are not password protected
- ✅ Verify latest file versions are being used
- ✅ Confirm TIES includes Month values
- ✅ Verify Cost Centers and PO Numbers are populated

### Configuration Review

Before generating a report:

- ✅ Validate Template settings
- ✅ Validate TIES settings
- ✅ Validate Forecast settings
- ✅ Confirm selected Cost Centers
- ✅ Review prior exceptions for recurring issues

### Output Validation

After report generation:

- ✅ Review Exception Summary
- ✅ Review high-priority exceptions
- ✅ Validate report totals
- ✅ Spot-check forecast values
- ✅ Spot-check actual values
- ✅ Save a copy of the final workbook

### Monthly Processing Checklist

- ✅ Latest Template Workbook
- ✅ Latest Vendor Forecast Files
- ✅ Latest TIES Extract
- ✅ Month column populated
- ✅ Configuration settings validated
- ✅ Prior exceptions reviewed
- ✅ Final report archived

---

## Common Workflows

### Monthly Report Generation

1. Obtain latest TIES extract.
2. Obtain updated forecast files.
3. Access the Financial Automation Portal.
4. Upload required files.
5. Review Cost Centers.
6. Verify configurations.
7. Generate report.
8. Review exceptions.
9. Download final workbook.

### Exception Review

1. Generate report.
2. Open Exception Summary.
3. Identify largest exception categories.
4. Review Exceptions Detail.
5. Investigate source data.
6. Correct data issues.
7. Re-run report.

### Configuration Maintenance

When file layouts change:

1. Review Template structure.
2. Review Forecast file structure.
3. Review TIES file structure.
4. Update configuration settings.
5. Test processing.
6. Validate output.
7. Document changes.

---

## Troubleshooting (Technical)

### Report Generation Fails

Verify:
- All required files are uploaded.
- Files are Excel (`.xlsx`) format.
- Required columns exist.
- Configuration settings are correct.
- Month column exists in the TIES file.

### Missing Data in Output

Verify:
- Transactions exist in the TIES file.
- Forecast values exist in vendor files.
- PO Numbers match across files.
- Cost Center selections are correct.
- Exceptions are reviewed.

### High Exception Volume

Review:
- Missing WBS records
- Missing PO records
- Duplicate WBS assignments
- Duplicate PO assignments
- TIES data completeness
- Configuration accuracy

---

*Version: 1.0 — Last Updated: July 2026*
