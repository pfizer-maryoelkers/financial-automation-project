from dataclasses import dataclass, field
from typing import Optional
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
    reclass_adjustments: dict[str, list] = field(default_factory=dict)
    # reclass_adjustments: {month_label: [(amount, description), ...]}
    legal_entity: Optional[str] = None
    country: Optional[str] = None
    vendor_name: Optional[str] = None
    gl_account: Optional[str] = None
    gross_po_value: Optional[float] = None
    req_title: Optional[str] = None
    real_wbs: Optional[str] = None  # Actual WBS from transactional file when PO is an ER
@dataclass
class WBSCode:
    wbs_code: str
    cost_center: Optional[str] = None
    pos: dict[str, PO] = field(default_factory=dict)
    charge_type: Optional[str] = None  # 'Capital' | 'Expense' — set by project pipeline
@dataclass
class CostCenter:
    cost_center_id: str
    wbs_codes: dict[str, WBSCode] = field(default_factory=dict)
@dataclass
class Project:
    """Top-level grouping for the CapEx / project pipeline.
    Groups by WBS root (e.g. 'CE-BTS21076') which is the prefix shared by
    all child WBS codes (CE-BTS21076-02-10, CE-BTS21076-02-EX, etc.)."""
    project_id: str          # e.g. 'CE-BTS21076'
    wbs_codes: dict[str, WBSCode] = field(default_factory=dict)

@dataclass
class P3ID:
    """Top-level grouping for the Projects pipeline (acts like Cost Center).
    Groups by P3 ID which contains WBS codes and POs."""
    p3_id: str
    wbs_codes: dict[str, WBSCode] = field(default_factory=dict)

    
# ---------------------------

## Building Exceptions Log

class ExceptionType(Enum):
    MISSING_WBS_AND_PO = "MISSING_WBS_AND_PO"
    MISSING_WBS = "MISSING_WBS"
    MISSING_PO = "MISSING_PO"
    DUPLICATE_PO = "DUPLICATE_PO"
    DUPLICATE_WBS = "DUPLICATE_WBS"
    RECLASS = "Reclass"
    UNMATCHED_TRANSACTION = "Unmatched Transaction"
    PO_NOT_ON_TEMPLATE = "PO Not on Template"
    UNMATCHED_P3 = "Unmatched P3 ID"
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
    vendor_name: Optional[str] = None
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
    
    def summary_by_month(self) -> dict:
        """Returns count of exceptions by month, with breakdowns by type and cost center"""
        result = {}
        for entry in self.entries:
            month = entry.month or 'Unknown'
            if month not in result:
                result[month] = {'total': 0, 'by_type': {}, 'by_cost_center': {}}
            
            result[month]['total'] += 1
            
            # By exception type
            exc_type = entry.exception_type.value
            if exc_type not in result[month]['by_type']:
                result[month]['by_type'][exc_type] = 0
            result[month]['by_type'][exc_type] += 1
            
            # By cost center
            cc = entry.cost_center or 'Unknown'
            if cc not in result[month]['by_cost_center']:
                result[month]['by_cost_center'][cc] = 0
            result[month]['by_cost_center'][cc] += 1
        
        return result