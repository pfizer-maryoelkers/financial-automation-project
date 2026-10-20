"""Front-end notices that depend only on the calendar."""

import calendar
import os
from datetime import date

# Day of the month from which the new actuals file may have been posted.
# Files are usually posted between the 25th and the last day of the month.
ACTUALS_WINDOW_START_DAY = 25


def actuals_file_notice(today: date | None = None) -> str | None:
    """Warning text when a new transactional detail (actuals) file may be out.

    The accounting period number follows the calendar month (AP07 = July), so in
    the posting window of October the file to look for is AP10.  Returns None
    outside the window (day 25 through the last day of the month).
    """
    if today is None:
        # Optional override for testing, e.g. ACTUALS_NOTICE_TEST_DATE=2026-10-26
        override = os.environ.get("ACTUALS_NOTICE_TEST_DATE", "").strip()
        try:
            today = date.fromisoformat(override) if override else date.today()
        except ValueError:
            today = date.today()
    # February always ends on the 28th (even in a leap year); every other month
    # ends on its own last day (30 or 31).
    last_day = 28 if today.month == 2 else calendar.monthrange(today.year, today.month)[1]
    if not (ACTUALS_WINDOW_START_DAY <= today.day <= last_day):
        return None
    return (
        f"The most up to date actuals file might be AP{today.month:02d} - "
        "please double check your team's activity."
    )
