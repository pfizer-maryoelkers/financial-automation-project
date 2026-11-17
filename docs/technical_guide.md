Technical Guide

# How Financial Automation Works, and How to Maintain It

How the OpEx and Project pipelines classify, place and write transactions, and how to change them safely.

Last updated November 2026

## System Overview: Two Pipelines, One Workflow

The tool fills a Pfizer financial spreadsheet template from three inputs: the transactional detail file (actuals), vendor forecast files, and, for OpEx only, an optional LE budget file. It writes the populated workbook plus exception reports for anything it could not place cleanly.

### OpEx Pipeline

Cost center → WBS → PO hierarchy. One main sheet with a cost-center list. Config: `configs/config_base.yaml`. Entry: `python -m cli.main`.

### Project (CapEx) Pipeline

P3 ID → WBS → PO hierarchy. One tab per P3 ID, cloned from the blank Project Template sheet. Config: `configs/config_project.yaml`. Entry: `python -m cli.project`.

`detect_template_type` picks the pipeline from the template, so `python -m cli.run` and the Streamlit app need no extra setting.

**Scope of this guide.** Rules below are tagged OpEx, Project or Both. Where the two rule sets differ, the difference is called out in a table. Rule numbers from the source rule sets are not used because they differ between the two documents.

## Input Files and Where They Come From

| File | Source | Posted by |
| --- | --- | --- |
| OpEx actuals (transactional detail) | Financial Management channel on Teams | Tanya Wilkins (usually via an announcement) |
| Project actuals (transactional detail) | Same Financial Management channel | Lucas Adam Feinstein (usually via an announcement) |
| IBM vendor forecast | [SSCS & IBM group chat on Teams](https://teams.microsoft.com/l/chat/19%3Ade9c0c2490324d27840556e2ef67ba88%40thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D) | Raju Parkala |

[Open the Financial Management channel in Teams](https://teams.microsoft.com/l/channel/19%3AasdbsnpPiIN5g7BB-eE9tsoJnridztD3XZZKF7bzhAY1%40thread.tacv2/Financial%20Management?groupId=1edad06c-5045-46ac-a8a1-5886354da421&tenantId=7a916015-20ae-4ad1-9170-eefd915e9272&ngc=true&allowXTenantAccess=true)

Sources for other vendors’ forecasts and the LE file are not documented yet. Add them here when confirmed. If the posters change, update this table.

## Templates: Structure and Required Setup

Based on the current blank templates. Both must be populated with their key list before a run.

Each button asks you to confirm before saving. Files are the blank templates as provided.

|  | OpEx_Template.xlsx | P3_Template.xlsx |
| --- | --- | --- |
| Main sheet | `OpEx Template` | `Project Template` (blank, cloned per P3 ID) |
| Required user input | Cost centers in column A under the `Cost Center` header (A1), above the `EXPENSE` row (A8). Blank file: A2 to A7 free. | P3 IDs in column A and tab names in column B of sheet `Insert P3 ID HERE` (headers `P3 ID`, `Tab Name` in row 1) |
| Anchors found | Row 8 `EXPENSE`, row 9 header (`Contact for PO/ER`, `PO Number/ER Number`, `PO Start Date`, `PO end date`, `Vendor Name`, `WBS`), row 11 `Previous Period Invoices`, row 12 `EXPENSE END`, row 13 `NOTE:-` | Same row layout on `Project Template`, plus `P3 ID` and `WBS Code` in A1:B1 and an `INVOICE TOTAL` subtotal in H8 |
| Failure if empty | “No cost centers were found” | “No P3 IDs were found” |

**Blank OpEx template is not fully blank.** Besides the visible sheet, the uploaded file contains four hidden sheets: `Petes P3` (a sample PO list with cost center 2329947, notes such as “Review with Pete”), `All POs as on 6-3-26`, `Exception_Data` and `Exception_Data1` (leftover exception rows). They look like remnants of an earlier run. Decide whether to delete them, since stale data, and potentially sensitive PO and vendor details, travel with every copy of the template. Also confirm the pipeline ignores them, particularly `Exception_Data`, which the tool itself writes.

## Run It Locally or in the App

**Hosted app:** [pfizer-financial-automation.streamlit.app](https://pfizer-financial-automation.streamlit.app/). The commands below are for running it locally.

Requirements: Python 3.10 or newer (tested on 3.11) and pip. Run everything from the project root.

```
git clone https://github.com/pfizer-maryoelkers/financial-automation-project
cd financial-automation-project
pip install -r requirements.txt

streamlit run app.py                           # UI (recommended)
python -m cli.run                              # auto-detect OpEx vs project
python -m cli.main                             # OpEx only
python -m cli.project                          # project only
python -m cli.run configs/config_project.yaml  # explicit config
uvicorn api.main:app --reload --port 8000      # API, see docs/README_API.md
```

Set file paths in the matching file under `configs/` before using the CLI.

## 1. How Transactions Are Classified Both

Classification is driven by the document-number prefix, confirmed against the line description, with fallbacks when the prefix does not decide.

| Order | Condition | Result |
| --- | --- | --- |
| 1 | Doc number starts with `5` | Actual |
| 1 | Doc number starts with `9` | Reclass |
| 1 | Doc number starts with `2` and amount > 0 | Accrual |
| 1 | Doc number starts with `2` and amount < 0 | Reversal |
| 2 | No prefix decision: description contains *reversal* / *accrual* / *reclass* / *invoice* or *vendor* | Reversal / Accrual / Reclass / Actual |
| 3 | Still undecided: amount < 0 | Reversal |
| 3 | Still undecided: amount ≥ 0 | Actual |
| Conflict | Prefix result and keyword result disagree | **Double Check** exception |

An unknown transaction type that cannot be classified at all is logged as **Unmatched Transaction**.

## 2. How POs Are Validated and Deduplicated Both

1. A valid PO must begin with `95`.
2. Read the PO column first.
3. If missing or invalid, search `CO Doc Line Item Txt` and extract a PO if one is found.
4. If still missing, log **Missing PO**.
5. Only the first occurrence of a duplicate PO is used. Later occurrences are logged as **Duplicate PO**.

## 3. How Expense Reports (ER) Are Handled Both

An ER number is searched for, in this order, in: GL Line Description, GL Transaction Description, Description, CO Doc Line Item Txt, ER Description. Example: `ER91389 - SECURITY - SECURITY`.

| Case | Behavior |
| --- | --- |
| ER found, **no** valid 95… PO | ER number becomes the PO. Type = ER. A standalone PO-style row is created with the ER number in the PO column. Works on blank and filled templates. |
| ER found, **valid** 95… PO | Transaction stays on that PO’s row. Amount is added to Actual. The Actual cell is filled light green and carries a cell note saying the transaction was an ER. |

## 4. Reclass: Logged, but Still Counted in Actual Both

- Log an exception of type **Reclass** for every reclass line.
- Do **not** write the Reclass row directly to the template.
- The reclass amount **still counts toward Actual**.
- If there is no real PO: create the pseudo-PO `<Document Number>RC` and convert the transaction to Actual.

**Doc discrepancy.** The README says “RC appended” with a space-style description; the rule sets say `<Document Number>RC` with no separator. Confirm against `src/pipeline` output and make the README match.

## 5. Month Placement: Where OpEx and Project Diverge Both

This is the biggest difference between the two pipelines.

| Transaction | OpEx | Project |
| --- | --- | --- |
| Actual | Period N → month N−1 (AP01 → prior-year December, e.g. Dec (25)) | No shift |
| ER | N−1 | No shift |
| Reversal | N−1 | No shift |
| Reclass | N−1 | No shift |
| Accrual | Direct, no shift (AP02 → Feb) | No shift |
| International PO (Country India / IN) | Actual, ER, Reversal, Reclass go back **2** months. Accrual unchanged. | No shift |
| Weekly vendors (Actual with a GL Posting Date) | If GL Posting Date month = Accounting Period month: post in that month, no shift. Otherwise use the Actual rule. | Not applicable (nothing shifts) |
| Forecast, India PO | Shifts back 1 month | No shift |

**Doc discrepancy.** The README states that “the project config uses the same rule” as OpEx, but the Project rule set says *Project Templates → No month shifting, applies to all transaction types*. Check `shift_months` in `configs/config_project.yaml` and update whichever side is stale. The README troubleshooting row “Wrong month for a transaction” (GL Posting Date, Country) is OpEx-only.

## 6. Project Only: WBS, Charge Type and P3 Matching Project

### WBS Root

Strip the suffix after the second hyphen group. `CE-BTS21076-02-EX` → root `CE-BTS21076`.

### Charge Type

WBS contains `EX` → Expense. Otherwise → Capital.

### P3 Matching

Match in this order: (1) WBS root, (2) Cost Center. If neither resolves, log **Unmatched P3 ID**. The P3 config sheet is any sheet whose name starts with “Insert” or “Enter” and contains “P3” (for example *Insert P3 ID HERE*). Each P3 ID goes in column A with its tab name in column B.

### On Template Column (Exceptions Tab)

| Value | Meaning |
| --- | --- |
| Yes | The line’s amount appears as its own figure on its PO’s row on its P3 ID’s tab. |
| No | It does not. That includes lines with no PO and lines whose amount was added into a total with other lines for the same PO, month and type. |

## 7. Amount Source and Aggregation Both

Amount source priority, first match wins:

1. `Amount - BER`
2. Any other BER column
3. The amount column named in config
4. `Amount - MAR`
5. `Amount`

Aggregate by **PO + Month + Type + Cost Center + WBS**, round to 2 decimals, and never write 0 or None. Zero and None cells stay blank.

## 8. Forecast Rules Both

- Forecast columns must end with `FTotal` (for example `... - FTotal`).
- Sum all forecast rows per PO.
- If a PO appears in multiple forecast files, use the **first file** only.
- OpEx only: India PO forecast shifts back 1 month.

## 9. Template Writing: Blanks Only by Default Both

- Default `overwrite = False`: never overwrite existing values, never overwrite formulas, only populate blanks.
- Always replace `#REF!`, even when overwrite is false.
- `overwrite = True` replaces everything.

## 10. Exception Catalog: What Each Flag Means Both

Exceptions appear on the **Exceptions** tab. Raw rows are on a hidden **Exception_Data** tab used by formulas.

| Type | Raised when | Pipeline |
| --- | --- | --- |
| Missing PO | No valid 95… PO after checking PO column and CO Doc Line Item Txt | Both |
| Duplicate PO | PO already used; later occurrences flagged | Both |
| Duplicate WBS | One WBS maps to multiple cost centers | Both |
| Reclass | Any reclass row (also counted in Actual) | Both |
| Unmatched Transaction | Unknown transaction type | Both |
| Not Worked on Template | Cost center mismatch with the template | Both |
| Unmatched P3 ID | P3 ID cannot be resolved from WBS root or cost center | Project (the OpEx rule set lists it, but it only applies to P3 resolution) |
| Double Check | Prefix and keyword classification conflict | Both |

## 11. Fixed Anchors: Never Rename or Move Both

Never rename or move these in templates. The readers locate them by exact text or name.

| Anchor | Where | Applies to |
| --- | --- | --- |
| Previous Period Invoices | Marker in column A | Both |
| Contact for PO/ER | Label | Both |
| EXPENSE END | Stop marker below the cost-center list | OpEx (README refers to the “EXPENSE” row; confirm exact text) |
| Sheet `ERP` (LE file) | Needs a Cost Center header in the top 60 rows | OpEx |
| P3 config sheet | Name starts with Insert or Enter and contains P3 | Project |
| Sheet `Project Template` | Blank template cloned per P3 ID | Project |

**Doc discrepancy.** The OpEx rule set lists the P3 Config and Project Template sheets as fixed anchors. Those only exist in project workbooks.

## 12. How Columns Are Matched

Columns are found by header text, not position (shared helper `src/common/col_utils.py`). Matching ignores capitalisation, extra spaces, line breaks and a trailing `*`. The header row can sit anywhere near the top of the sheet.

Limits: a renamed header (“Cost Center\*” → “CC”) will not match. Add the new name to the reader’s alias list. The anchors in section 11 must stay as they are.

A transactional sheet is only valid if it has a period column (Accounting Period, Fiscal Year/Period, Period or Month), a PO column and an amount column. Otherwise the run fails with “No valid sheets found”.

## 13. What the Output Workbook Contains

| Sheet | Contents |
| --- | --- |
| Populated template / one tab per P3 ID | Values by month, ER rows, light-green ER notes |
| Budget (OpEx with LE file) | UID, cost center, description, project code, LE0 and every LE column found, with totals |
| Exception_Data (hidden) | Raw exception rows for formulas |
| Exceptions | The exception report |

## 14. Code Map and Configuration

```
app.py                    Streamlit entry
configs/                  config_base.yaml, config_project.yaml, config_streamlit.yaml
webapp/                   streamlit_backend, streamlit_config, pfizer_theme
api/                      FastAPI: main, backend, config
cli/                      run (auto-detect), main (OpEx), project (P3)
src/models.py             Data models and exception types
src/readers/              transactional_detail_reader, forecast_reader,
                          template_reader, project_template_reader
src/writers/              template_writer (workbook, Exceptions, Budget)
src/pipeline/             utils (OpEx hierarchy), project_utils (project hierarchy)
src/common/               col_utils, notices (banner), fast_excel (cache)
tools/                    One-off maintenance scripts
```

Key config sections: `template` (header row, PO column, stop marker), `forecast_reader`, `transactional_detail_reader` (colmap, valid_types, required_cols), `le_file`, `template_writer`.

## 15. Making Changes Safely

| Change | Where to look | Then |
| --- | --- | --- |
| Header renamed in a source file | Reader alias list (`src/common/col_utils.py` helper, per-reader aliases) | Add the new name; re-run a known file |
| New transaction prefix or keyword | Classification logic in `src/pipeline` and `valid_types` in config | Update the table in section 1 and test the Double Check path |
| Change month shifting | `shift_months` and month-placement logic; OpEx vs project config | Update section 5 and the README |
| New country needing a 2-month shift | International PO check (currently Country India / IN) | Update both shift and forecast rules |
| New amount column | `transactional_detail_reader` colmap; amount priority list | Keep the priority order documented in section 7 |
| Template layout change | Anchors (section 11) and `template` config | Never rename anchors; test a blank and a filled template |
| New exception type | `src/models.py` and the Exceptions writer | Add to section 10 and the Exceptions tab docs |

### Regression Checklist After Any Change

- Run one OpEx template and one Project template end to end (actuals and forecast).
- Check an AP01 row lands in prior-year December (OpEx).
- Check an India PO shifts two months for actuals and one month for forecast (OpEx), and does not shift in Project.
- Check a weekly vendor row where GL Posting month equals Accounting Period month.
- Check an ER with and without a 95… PO, and a reclass with and without a PO.
- Confirm no existing value or formula was overwritten and `#REF!` cells were replaced.
- Confirm every unplaced transaction appears on the Exceptions tab.

## 16. Deployment and Troubleshooting

Production app: [pfizer-financial-automation.streamlit.app](https://pfizer-financial-automation.streamlit.app/) (Streamlit Community Cloud). Confirm its sharing setting is restricted to approved viewers and that IT security has signed off, since uploads contain financial transaction data.

- Entry point `app.py` at the repo root. Keep `static/` and `.streamlit/` with the code.
- A hosted build needs only streamlit, pandas, openpyxl, PyYAML, numpy and python-dateutil. Add fastapi, uvicorn and python-multipart only for the API.
- **Data sensitivity:** uploads contain financial transaction data. Do not host publicly without IT security approval. The repo also holds Pfizer brand fonts and logos.
- No persistence: uploads and outputs are in a temp folder. Download the workbook right after each run.
- Large files use `src/common/fast_excel.py`; parsed files are cached in `financial_automation_cache` in the system temp folder (cleared after 7 days, keyed on file contents).
- Upload limit is 200 MB per file. Raise `server.maxUploadSize` in `.streamlit/config.toml` if needed.
- Month-end banner shows from the 25th. Preview with `ACTUALS_NOTICE_TEST_DATE=2026-10-26`.

| Symptom | Check |
| --- | --- |
| “No valid sheets found” | Needs a period column, a PO column and an amount column |
| A column is not found | Header text. Spacing, case and * don’t matter, a rename does |
| LE file not read | ERP tab with a Cost Center header in the top 60 rows |
| Logo or fonts missing | Restart, Ctrl+F5, confirm `static/` deployed |
| Generate Report greyed out | Upload a template first |
| “No cost centers were found” | OpEx: list them in column A under “Cost Center”, above the EXPENSE row |
| “No P3 IDs were found” | Project: P3 IDs in column A, tab name in column B on the P3 config sheet |
| Wrong month (OpEx) | Check GL Posting Date, Country and period against section 5 |

Internal Pfizer project. Maintained by Mary Oelkers. Source: the [project’s GitHub repository](https://github.com/pfizer-maryoelkers/financial-automation-project) and the documented OpEx and Project processing rules.
