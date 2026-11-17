"""Turn pipeline exceptions into specific, actionable messages for the Streamlit UI."""

# (substring of the exception text, what went wrong, how to fix it) — first match wins.
_KNOWN_ERRORS = (
    ("no valid sheets found",
     "The actuals file could not be read.",
     "Your actuals file needs a period column, a PO column and an amount column."),
    ("no cost centers were found",
     "No cost centers were found in the OpEx template.",
     "List cost centers in column A under “Cost Center”, above the EXPENSE row."),
    ("no p3 ids were found",
     "No P3 IDs were found in the Project template.",
     "On the P3 config sheet, put each P3 ID in column A and its tab name in column B."),
    ("expected a sheet named 'erp'",
     "The LE file could not be read, so the Budget tab and UIDs are missing.",
     "The LE file needs an ERP tab with a Cost Center header near the top."),
    ("could not find a 'cost center' header",
     "The LE file could not be read, so the Budget tab and UIDs are missing.",
     "The LE file needs an ERP tab with a Cost Center header near the top."),
    ("currently open in microsoft excel",
     "The output file could not be saved.",
     "Close the output file in Excel and run again."),
    ("is not a zip file",
     "One of the uploaded files is not a valid .xlsx workbook.",
     "Open the file in Excel and use Save As → Excel Workbook (.xlsx), then upload it again."),
)


def explain_error(exc: BaseException) -> tuple[str, str | None]:
    """Return (what went wrong, how to fix it or None) for an exception."""
    text = str(exc)
    low = text.lower()
    for needle, problem, fix in _KNOWN_ERRORS:
        if needle in low:
            return problem, fix
    first_line = text.strip().splitlines()[0] if text.strip() else type(exc).__name__
    return first_line, None


# Rows for the "Troubleshooting" panel: (what you see, what to do).
TROUBLESHOOTING = (
    ("Generate Report is grey", "Upload a template first."),
    ("“No valid sheets found”",
     "Your actuals file needs a period column, a PO column and an amount column."),
    ("“No cost centers were found” (OpEx)",
     "List cost centers in column A under “Cost Center”, above the EXPENSE row."),
    ("“No P3 IDs were found” (Project)",
     "On the P3 config sheet, put each P3 ID in column A and its tab name in column B."),
    ("A number is in the wrong month",
     "Check the period, the GL posting date and whether the PO is an India PO. "
     "Remember Project templates do not shift."),
    ("Budget tab missing or LE not read",
     "The LE file needs an ERP tab with a Cost Center header near the top."),
    ("Logo or fonts missing", "Restart the app and refresh with Ctrl+F5."),
    ("First run is slow", "Normal for big files. Running the same file again is much faster."),
)
