from dataclasses import dataclass, field
from typing import Optional
from enum import Enum
import pandas as pd

@dataclass
class MonthlyMetrics:
    forecast: float = 0.0
    actual: float = 0.0
    accrual: float = 0.0
    accrual_reversal: float = 0.0
@dataclass
class PO:
    po_number: str
    monthly_data: dict[str, MonthlyMetrics] = field(default_factory=dict)
@dataclass
class WBSCode:
    wbs_code: str
    cost_center: Optional[str] = None
    pos: dict[str, PO] = field(default_factory=dict)
@dataclass
class CostCenter:
    cost_center_id: str
    wbs_codes: dict[str, WBSCode] = field(default_factory=dict)

    
# ---------------------------

## Building Exceptions Log

class ExceptionType(Enum):
    MISSING_WBS = "MISSING_WBS"
    MISSING_PO = "MISSING_PO"
    MISSING_FORECAST = "MISSING_FORECAST"
    DUPLICATE_PO = "DUPLICATE_PO"
@dataclass
class ExceptionEntry:
    exception_type: ExceptionType
    row_index: Optional[int] = None
    po: Optional[str] = None
    wbs: Optional[str] = None
    cost_center: Optional[str] = None
    detail: Optional[str] = None
@dataclass
class ExceptionLog:
    entries: list[ExceptionEntry] = field(default_factory=list)
    def log(self, exception_type: ExceptionType, **kwargs):
        self.entries.append(ExceptionEntry(exception_type=exception_type, **kwargs))