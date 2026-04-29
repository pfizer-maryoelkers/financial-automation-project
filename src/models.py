from dataclasses import dataclass, field
from typing import Optional
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

    