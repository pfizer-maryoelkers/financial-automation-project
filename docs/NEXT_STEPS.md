# Recommended Next Development Steps

This document outlines some next steps for this project that would improve functionality of the automation system. These next steps are ranked in terms of impact and complexity.

## 1. Normalized Starting Template - High Impact, Low Complexity

Using new template (located at data/templates/financial_template_v1.xlsx) as the starting point for all new spreadsheets. For new orgs, use this same file as the starting point. New template is mostly finished, although would still benefit from further enhancements.

Key Improvements:
* Aestetic Improvements
* Simplifies configurations. Can now use standard parameters for all template configs - removing the need to update these.

## 2. Improved Transactional Detail Configurations - Medium Impact, Medium Complexity

A few steps can be taken to improve the transactional detail configurations. Some ideas may include:
* Using 'Accounting Period' column instead of 'Month'. This handles the case where the transactional detail file does not contain a 'Month' column.
* Adding better error handling. Right now, system throws a generic error when it cannot read the transacional detail file. It would be worth adding more descriptive errors - for example it would throw an error that explicitly says "Following columns are missing in Transactional Detail File: Month, Cost Center, ..." if column map is not properly configured.

Key Improvements:
* Simplified Transactional Detail configurations.
* Easier debugging

## 3. Improved ER support - High Complexity, High Impact

Right now, ERs are handled the same way at POs. They will be caught as long as they have a cost center and WBS code, however they are not distinguished from regular PO numbers. 

In order to distinguish ER, must look at "Description" column and parse this text, looking for things like "ER". Can possibly add a boolean data field called "ER" that is either True or False depending on the row.

After rows are distinguished, must also decide how to write this data. May have to speak with Kinjal and Kerry about how ERs are traditionally handled, then update template writer. Do they go on main sheet, or exceptions sheet?

Key Improvements:
* Handles more cases, reduces manual efforts in accounting.

## 4. Enhanced Frontend - High Complexity, High Impact

Improve streamlit UI, or move to full stack application if necessary.

If moving to a full stack application, this may require refactoring or reworking how configurations are handled. Currently not a clear path forward here so it would require some further investigation before proceeding.

Key Improvements:
* Better user experience