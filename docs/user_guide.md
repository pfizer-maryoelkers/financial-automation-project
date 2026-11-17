User Guide

# Turn Actuals and Forecasts Into Finished Financial Templates

A plain-language guide to how the tool fills your OpEx and Project templates, and what to check afterwards.

Last updated November 2026

## What You Give It, What You Get Back

It fills in your financial spreadsheet template automatically. You give it your template, last month’s actuals and the new forecast. It puts every number in the right row and month, and lists anything it was unsure about so you can check it.

### You Provide

- Your **template** (OpEx or Project)
- The **actuals file** (transactional detail)
- The **forecast file(s)** from vendors
- OpEx only, optional: the **LE file**

### You Get Back

- Your template, **filled in**
- An **Exceptions** tab listing what needs a human look
- OpEx with an LE file: a **Budget** tab

## Get Your Files

### OpEx Actuals OpEx

Posted in the **Financial Management** channel on Microsoft Teams. An announcement is usually posted by **Tanya Wilkins**.

### Project Actuals Project

Posted in the same **Financial Management** channel. An announcement is usually posted by **Lucas Adam Feinstein**.

### IBM Forecast

Found in the **SSCS & IBM** group chat on Microsoft Teams, posted by **Raju Parkala**.

[Open the SSCS & IBM chat](https://teams.microsoft.com/l/chat/19%3Ade9c0c2490324d27840556e2ef67ba88%40thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D)

[Open the Financial Management channel in Teams](https://teams.microsoft.com/l/channel/19%3AasdbsnpPiIN5g7BB-eE9tsoJnridztD3XZZKF7bzhAY1%40thread.tacv2/Financial%20Management?groupId=1edad06c-5045-46ac-a8a1-5886354da421&tenantId=7a916015-20ae-4ad1-9170-eefd915e9272&ngc=true&allowXTenantAccess=true)

Look for the newest announcement before you start, especially in the last week of the month. Where other vendors’ forecasts and the LE file come from is not documented here yet.

## Set Up Your Template

The blank templates do not know which cost centers or projects are yours. Fill these in **before** you upload, or the tool has nothing to build on.

Each button asks you to confirm before saving. Files are the blank templates as provided.

### OpEx: Add Your Cost Centers OpEx

1. Open **OpEx_Template.xlsx** (sheet *OpEx Template*).
2. Find the **Cost Center** header in cell A1.
3. Type one cost center per row in **column A**, directly under that header, in the rows above the **EXPENSE** row (A2 to A7 in the blank file).
4. Leave everything else as is.

### Project: Add Your P3 IDs Project

1. Open **P3_Template.xlsx** and go to the sheet *Insert P3 ID HERE*.
2. Under **P3 ID** in column A, list one P3 ID per row.
3. In **column B** (*Tab Name*), type the tab name you want for that project.
4. Leave the *Project Template* sheet blank. The tool copies it once for each P3 ID.

**If you skip this step** the app will stop with “No cost centers were found” (OpEx) or “No P3 IDs were found” (Project). Do not rename the sheets or move the markers listed under “What the tool changes, and what it leaves alone”.

## Run the Tool in Six Steps

[Open the App](https://pfizer-financial-automation.streamlit.app/)

1. Open the app. You will see the Pfizer-branded upload page.
2. Upload your **template**. A banner tells you whether it is OpEx or Project. You do not have to choose.
3. Upload the **actuals file** and any **forecast files**. For OpEx, add the **LE file** if you want the Budget tab.
4. Click **Generate Report**. The button stays grey until a template is uploaded.
5. Download the workbook **right away**. The app does not save your files.
6. Open the **Exceptions** tab and work through anything listed.

**Month-end reminder.** From the 25th through the end of the month the app shows a banner saying the most up-to-date actuals file might be the next AP period. Double-check you are using the newest file.

**Keep it private.** These files hold financial transaction data. Only use the approved version of the app.

## OpEx vs Project: What Changes

|  | OpEx | Project (CapEx) |
| --- | --- | --- |
| Organized by | Cost center, then WBS, then PO | P3 ID, then WBS, then PO |
| Template | One main sheet listing your cost centers | A sheet where you list your P3 IDs, plus a blank Project Template sheet |
| Result | Your template, filled in | One tab per P3 ID |
| LE file | Optional. Adds UIDs and a Budget tab | Not used |
| Months | Most actuals land **one month earlier** than their period. India POs land two months earlier | **Nothing is shifted.** The month is the month |
| Extra step | None | The tool works out the project (P3 ID) from the WBS code or cost center and tells you if it cannot |

## How Transactions Get Sorted

Every line in the actuals file is put into one bucket, based mostly on the first digit of its document number.

| Starts with | What it is | Bucket |
| --- | --- | --- |
| 5 | Vendor invoice | **Actual** |
| 9 | Reclass (moved between accounts) | **Reclass** (counted in Actual too) |
| 2, amount positive | An estimate of cost not yet invoiced | **Accrual** |
| 2, amount negative | Undoing a prior accrual | **Reversal** |

If the number does not settle it, the tool reads the description (words like invoice, accrual, reversal or reclass). If that still does not settle it, a negative amount is treated as a reversal and anything else as an actual. If the digit and the words disagree, the line goes on the Exceptions tab as **Double Check** for you to decide.

## How POs and Expense Reports Are Handled

A real PO number starts with **95**. The tool looks in the PO column first, then in the line text. If it cannot find one, the line is flagged **Missing PO**. If the same PO shows up again, only the first one is used and the repeats are flagged **Duplicate PO**.

### Expense Reports (ER)

If a line mentions an expense report number, like `ER91389`, the tool treats it in one of two ways.

| Situation | What you will see |
| --- | --- |
| ER with **no** 95 PO | A new row of its own, with the ER number in the PO column |
| ER with a real 95 PO | The amount is added to that PO’s Actual. The cell is light green with a note saying it was an ER |

### Reclass

Reclass lines are listed on the Exceptions tab so you can see them, but the amount still counts in Actual. If a reclass has no real PO, the tool gives it a made-up PO made of the document number plus `RC` and treats it as an actual.

## Where Each Number Lands by Month

Examples of where actuals and forecast numbers land.

| Example | OpEx lands in | Project lands in |
| --- | --- | --- |
| Domestic actual, invoice or reclass, period 09 | **Aug** (one month back) | **Sep** (no shift) |
| Accrual, period 09 | **Sep** (no shift) | **Sep** |
| Domestic actual, period 01 | **December of the prior year** | **Jan** |
| India PO actual, period 09 | **Jul** (two months back) | **Sep** |
| India PO forecast | One month earlier than listed | As listed |
| Weekly vendor, posting month matches period month | That month, no shift | That month |

**Why the shift?** For OpEx, a period’s invoices are generally booked for the month before. Weekly vendors whose GL posting date falls in the same month as the period are the exception and stay put.

## What the Tool Changes, and What It Leaves Alone

- It **only fills blank cells**. It does not change numbers you already typed and does not touch formulas.
- Broken `#REF!` cells are always replaced.
- Zero amounts are left blank.
- Amounts for the same PO, month, type, cost center and WBS are added together into one number.
- Forecasts are summed per PO. If a PO is in more than one forecast file, only the first file is used.

**Please do not rename or move** these template items: *Previous Period Invoices*, *Contact for PO/ER* and the *EXPENSE END* marker (OpEx). Project templates also depend on the *P3 config* sheet and the *Project Template* sheet. The ERP sheet in the LE file must keep its name too. Moving columns is fine. Renaming column headers is not.

## Your To-Do List: The Exceptions Tab

Anything the tool could not place cleanly is listed here. This is your to-do list before sending the file on.

| Exception | What it means | What to do |
| --- | --- | --- |
| Missing PO | No valid 95 PO found | Find the PO and add it to the source data or the template |
| Duplicate PO | The same PO appeared again | Check whether it is a true duplicate; the first one was used |
| Duplicate WBS | One WBS code is tied to more than one cost center | Confirm the right cost center |
| Reclass | A reclass line (still counted in Actual) | Review and confirm it is correct |
| Unmatched Transaction | The tool could not tell what type of line it is | Classify by hand |
| Not Worked on Template | Cost center does not match your template | Check that cost center belongs on your template |
| Unmatched P3 ID Project | The project could not be figured out | Check the WBS code and add the P3 ID to the config sheet |
| Double Check | Document number and description disagree about the type | Decide which is right and enter by hand |

**Project Exceptions tab, “On Template” column:** *Yes* means the line’s amount shows as its own figure on its PO row. *No* means it does not, for example because it had no PO or was added together with other lines for the same PO, month and type.

## Fix Common Problems

| What you see | Try this |
| --- | --- |
| Generate Report is grey | Upload a template first |
| “No valid sheets found” | Your actuals file needs a period column, a PO column and an amount column |
| “No cost centers were found” (OpEx) | List cost centers in column A under “Cost Center”, above the EXPENSE row |
| “No P3 IDs were found” (Project) | On the P3 config sheet, put each P3 ID in column A and its tab name in column B |
| A number is in the wrong month | Check the period, the GL posting date and whether the PO is an India PO. Remember Project templates do not shift |
| Budget tab missing or LE not read | The LE file needs an ERP tab with a Cost Center header near the top |
| Logo or fonts missing | Restart the app and refresh with Ctrl+F5 |
| First run is slow | Normal for big files. Running the same file again is much faster |

Still stuck? Contact Mary Oelkers, who maintains the tool.

## Final Check Before You Send

- Downloaded the workbook straight after the run.
- Used the latest actuals file and the right forecast.
- Worked through every line on the Exceptions tab.
- Checked the light-green ER cells.
- Spot-checked a few POs for the right month.

Internal Pfizer project. Maintained by Mary Oelkers. Source: the [project’s GitHub repository](https://github.com/pfizer-maryoelkers/financial-automation-project) and the documented OpEx and Project processing rules.
