Governance

# Governance for Financial Automation: Ownership, Controls and Change

Who owns the tool, which items are controlled, how changes are approved, and how financial data is protected.

Draft for review · Last updated November 2026

## Purpose and Scope

This document sets out who owns the financial automation tool, which items are controlled, how changes are approved and released, and how data is protected. It covers the app, the code and configuration, the OpEx and Project processing rules, the blank templates, the input files and the output workbooks.

**How to read this draft.** Items tagged Confirmed come from the README, the rule sets or what the project team has told us. Items tagged Proposed are recommended controls that still need an owner’s approval. Nothing tagged Proposed should be quoted as Pfizer policy until it is confirmed.

### Principles

- The documented rules are the single source of truth. Code, configuration and both guides must agree with them.
- The tool never silently overwrites. By default it fills blanks only, leaves formulas alone and replaces `#REF!`.
- Anything it cannot place cleanly is shown on the Exceptions tab, never dropped.
- People review exceptions before a workbook is sent on.
- Financial transaction data is handled only in approved environments.

## Roles and Responsibilities

| Role | Who | Responsibilities | Status |
| --- | --- | --- | --- |
| Tool owner and maintainer | Mary Oelkers | Owns the code, configuration, rules documentation and guides. Approves and releases changes. Answers user questions. | Confirmed |
| OpEx actuals provider | Tanya Wilkins | Posts the OpEx actuals file, usually with an announcement in the Financial Management channel on Teams. | Confirmed |
| Project actuals provider | Lucas Adam Feinstein | Posts the Project actuals file, usually with an announcement in the same channel. | Confirmed |
| IBM forecast provider | Raju Parkala | Posts the IBM forecast in the SSCS & IBM group chat on Teams. | Confirmed |
| Users | Finance teams that run the tool | Prepare templates, run the tool, work through exceptions, store outputs appropriately. | Confirmed |
| Rule approver | To be named | Approves any change to classification, month placement, amount or forecast rules before release. | Proposed |
| Security approver | IT security (contact to be named) | Approves hosting and access for an app that handles financial transaction data. The README requires this approval. | Confirmed |
| Backup maintainer | To be named | Can run, fix and release the tool if the owner is unavailable. | Proposed |

## Controlled Items

Each item below has one owner and one place where its approved version lives.

| Item | Approved Version Lives In | Change Trigger |
| --- | --- | --- |
| OpEx and Project processing rules | The documented rule sets, mirrored in the technical guide | A change to classification, month placement, amount source, forecast or exception handling |
| Code | GitHub repository `pfizer-maryoelkers/financial-automation-project` | Bug fix, new rule, new column alias |
| Configuration | `configs/config_base.yaml`, `config_project.yaml`, `config_streamlit.yaml` | New file layout, new column name, shift settings |
| Blank templates | OpEx and Project templates (downloadable from the guides) | Layout change, new anchors, clean-up of hidden sheets |
| Hosted app | [pfizer-financial-automation.streamlit.app](https://pfizer-financial-automation.streamlit.app/) | Release, access change, hosting change |
| Guides | User guide and technical guide | Any change to the items above |
| Input data sources | Teams: Financial Management channel, SSCS & IBM chat | A provider or location changes |

## Data Protection and Hosting

### What Is Confirmed

- Uploads contain financial transaction data. The README says not to host the app on a public service without IT security approval.
- The app keeps no uploads or outputs. Files live in a temporary folder, so users download the workbook right after each run.
- Parsed transaction files are cached in the system temp folder (`financial_automation_cache`) for up to 7 days. The cache is keyed on file contents.
- The repository also holds Pfizer brand fonts and logos.
- The production app runs at a `streamlit.app` address, which is Streamlit Community Cloud.

### Actions Needed

| Action | Why | Status |
| --- | --- | --- |
| Confirm the app’s sharing setting is restricted to approved viewers and record the IT security approval | Community Cloud apps are public by default unless restricted. Approval is required by the README. | Proposed |
| Remove or justify the four hidden sheets in the OpEx template (`Petes P3`, `All POs as on 6-3-26`, `Exception_Data`, `Exception_Data1`) | They hold PO numbers, vendor names and a cost center, and travel with every copy of the template. | Proposed |
| Define where finished workbooks are stored and for how long | The app retains nothing, so retention is the user’s responsibility. | Proposed |
| Keep the repository private and limit write access to named maintainers | The repo holds brand assets and sample data folders. | Proposed |
| Do not commit real transaction data, extracts or cache files to the repository | Sample inputs under `data/` should be anonymised or removed. | Proposed |

## Change Management

Different changes need different levels of control.

| Change Type | Example | Minimum Control |
| --- | --- | --- |
| Rule change | Shift months for a new country; change amount priority | Written request, rule approver sign-off, update the rule sets, config, both guides and the regression checklist, test both pipelines |
| Configuration change | New column alias, new header name | Owner review, test with a known file, note it in the change log |
| Template change | New layout, moved anchors | Never rename or move fixed anchors. Test a blank and a filled template. Re-issue the template and update the downloads in the guides |
| Code fix | Bug in placement or exceptions | Branch, test, owner review, release |
| New exception type | A new flag on the Exceptions tab | Update the model, writer, both guides and the user to-do table |
| New input source | A new vendor forecast | Add provider and location to both guides |

### Standard Steps

1. Raise the request with the reason, an example transaction and the expected result.
2. Assess impact on OpEx, Project or both, and on templates and guides.
3. Make the change on a branch, never directly on the released version.
4. Run the regression checklist in the technical guide for one OpEx and one Project template.
5. Obtain approval. Rule changes need the rule approver as well as the owner.
6. Release, then tell users what changed and from which period it applies.
7. Update the README “Last updated” date and both guides.

The steps above are Proposed and should be confirmed by the tool owner.

## Version Control and Documentation Accuracy

Use GitHub history as the audit trail. Tag each release, keep a short change log and record which rule version a given month’s run used.

### Open Documentation Gaps

These mismatches were found while writing the guides. Resolve each by fixing the side that is out of date, then record the decision.

| Gap | Detail |
| --- | --- |
| Project month shifting | The README says the project config uses the OpEx rule, but the Project rules say no month shifting. Check `shift_months` in `config_project.yaml`. |
| Reclass pseudo-PO format | The README describes “RC appended”. The rules say `<Document Number>RC`. Confirm the exact format. |
| Fixed anchors by template | The OpEx rules list the P3 Config and Project Template sheets, which exist only in project workbooks. |
| EXPENSE marker text | The README refers to an “EXPENSE” row. The rules refer to “EXPENSE END”. Confirm which text readers use. |
| Exception names | Missing PO, Duplicate PO and Duplicate WBS were renamed on the Exceptions tab. Confirm the README and code constants match. |
| Single maintainer | One named maintainer is a continuity risk. Name a backup. |

## Monthly Operating Cadence

| Stage | Control | Owner |
| --- | --- | --- |
| Before | Find the newest actuals announcement. Download the right actuals and forecast files. Confirm the template has cost centers (OpEx) or P3 IDs (Project) filled in. | User |
| Run | Upload files, confirm the OpEx or Project banner, generate the report and download the workbook straight away. | User |
| After | Work through every line on the Exceptions tab. Check light-green ER cells. Spot-check a few POs for the right month. | User |
| Month end | Heed the banner from the 25th about the newest actuals file. | User |
| Release | Run the regression checklist before any change goes live. | Tool owner |

The cadence reflects the guides. Sign-off on exception review before a workbook is sent on is Proposed.

## Exception Ownership

Every exception needs a human decision. Proposed resolution owners:

| Exception | Typical Resolver | Escalate When |
| --- | --- | --- |
| Missing PO | Template or cost center owner | The PO cannot be found in the source data |
| Duplicate PO | Template owner | The repeat looks like a true second charge |
| Duplicate WBS | Cost center owner | One WBS is tied to more than one cost center |
| Reclass | Template owner | The reclass is unexpected or large |
| Unmatched Transaction | Template owner | The type cannot be classified by hand |
| Not Worked on Template | Template owner | The cost center should be on the template |
| Unmatched P3 ID | Project lead | The P3 ID or WBS root is missing from the config sheet |
| Double Check | Template owner | Document number and description disagree |

If the same exception recurs month after month, treat it as a rule or template gap and raise a change request. Owners and escalation thresholds are Proposed.

## Built-In Controls

| Control | What It Protects | Evidence |
| --- | --- | --- |
| Fill blanks only by default | Manually entered numbers and formulas | Overwrite setting off unless chosen |
| Exceptions tab and hidden Exception_Data | Completeness: nothing unplaced goes unseen | Exceptions tab in every output |
| Double Check flag | Classification conflicts between prefix and description | Exceptions tab |
| Duplicate PO and Duplicate WBS flags | Double counting and wrong ownership | Exceptions tab |
| Light-green ER cells with a note | Expense report amounts on real POs | Cell colour and note |
| Fixed anchors and input validation messages | Template integrity | Errors such as “No valid sheets found” |
| Cache keyed on file contents | Stale parsed data | An edited file is always re-read |

### Incident Handling

1. Stop using the affected output and keep the input files.
2. Tell the tool owner what was wrong, with a sample transaction and the month.
3. Owner finds the cause, fixes it through change management and reruns the period.
4. Users replace the affected workbooks and the change is logged.

Response times and who must be notified are Proposed.

## Review Cadence

| Review | Frequency | Scope |
| --- | --- | --- |
| Rules | Quarterly, and whenever the source data format changes | Check rules, config and guides still agree |
| Access | Twice a year | App sharing, repository access, Teams provider list |
| Documentation | After every release | Guides, README date, download links |
| Data | Annually | Hosting approval, retention, template contents |

This schedule is Proposed.

## Decisions Needed

1. Who approves rule changes, and who is the backup maintainer?
2. Is the Streamlit app restricted, and where is the IT security approval recorded?
3. Where are finished workbooks stored and for how long?
4. Should the OpEx template’s four hidden sheets be removed?
5. Does exception review need a recorded sign-off?
6. Where do the other vendor forecasts and the LE file come from, and who posts them?
7. Which side of each documentation gap is correct?

Internal Pfizer project. Maintained by Mary Oelkers. Source: the [project’s GitHub repository](https://github.com/pfizer-maryoelkers/financial-automation-project) and the documented OpEx and Project processing rules.
