from dataclasses import dataclass, field
from typing import Counter, Optional
from enum import Enum
import pandas as pd
from collections import Counter

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
    MISSING_WBS_AND_PO = "MISSING_WBS_AND_PO"
    MISSING_WBS = "MISSING_WBS"
    MISSING_PO = "MISSING_PO"
    DUPLICATE_PO = "DUPLICATE_PO"
    DUPLICATE_WBS = "DUPLICATE_WBS"
@dataclass
class ExceptionEntry:
    exception_type: ExceptionType
    row_index: Optional[int] = None
    po: Optional[str] = None
    wbs: Optional[str] = None
    cost_center: Optional[str] = None
    month: Optional[str] = None
    amount: Optional[float] = None
    transaction_type: Optional[str] = None
    source_row_data: Optional[dict] = None
@dataclass
class ExceptionLog:
    entries: list[ExceptionEntry] = field(default_factory=list)
    def log(self, exception_type: ExceptionType, **kwargs):
        self.entries.append(ExceptionEntry(exception_type=exception_type, **kwargs))

    def summary(self):
        counts = Counter(e.exception_type.value for e in self.entries)
        if counts:
            for exc_type, count in counts.items():
                print(f"  {exc_type}: {count}")
        else:
            print("  No exceptions logged.")
    
    def summary_by_type(self) -> dict:
        """Returns count of exceptions by type"""
        counts = Counter(e.exception_type.value for e in self.entries)
        total = len(self.entries)
        return {
            'counts': dict(counts),
            'total': total,
            'percentages': {k: (v / total * 100) if total > 0 else 0
                          for k, v in counts.items()}
        }
    
    def summary_by_cost_center(self) -> dict:
        """Returns count of exceptions by cost center and type"""
        result = {}
        for entry in self.entries:
            cc = entry.cost_center or 'Unknown'
            exc_type = entry.exception_type.value
            
            if cc not in result:
                result[cc] = {'total': 0, 'by_type': {}}
            
            result[cc]['total'] += 1
            if exc_type not in result[cc]['by_type']:
                result[cc]['by_type'][exc_type] = 0
            result[cc]['by_type'][exc_type] += 1
        
        return result