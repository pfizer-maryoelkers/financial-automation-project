"""
Configuration management for the FastAPI layer.
Mirrors streamlit_config.py but without any Streamlit dependency.
"""

from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Any
from pathlib import Path
import yaml


@dataclass
class TemplateConfig:
    header_row: int = 16
    po_col: str = "B"
    po_stop_marker: str = "Previous Period Invoices"
    cost_center_col: str = "A"
    cost_center_start_row: int = 9
    cost_center_end_row: Optional[int] = None


@dataclass
class ForecastConfig:
    po_col: str = "PO #"


@dataclass
class TransactionalConfig:
    required_cols: List[str] = field(default_factory=lambda: [
        "PO Number", "Accounting Period", "GL Transaction Amount"
    ])
    valid_types: List[str] = field(default_factory=lambda: [
        "Actual", "Accrual", "Reversal", "Reclass", "ER"
    ])
    colmap: Dict[str, str] = field(default_factory=lambda: {
        'po': 'PO Number',
        'month': 'Accounting Period',  # YYYYMM format e.g. 202601 → Jan
        'amount': 'GL BER Corp Amount',
        'classifier': 'AP Voucher Number',
        'cost_center': 'CC ID',
        'wbs': 'WBS Element',
        'type': 'Type'
    })


@dataclass
class WriterConfig:
    output_path: str = "template_output.xlsx"
    overwrite: bool = True   # API always overwrites into the temp dir
    dec_acc_reversal_col: str = "N"
    forecast_source_cols: List[str] = field(default_factory=lambda: [
        "PO #",
        "Jan 2026 - FTotal", "Feb 2026 - FTotal", "March 2026 - FTotal",
        "April 2026 - FTotal", "May 2026 - FTotal", "June 2026 - FTotal",
        "July 2026 - FTotal", "Aug 2026 - FTotal", "Sep 2026 - FTotal",
        "Oct 2026 - FTotal", "Nov 2026 - FTotal", "Dec 2026 - FTotal"
    ])
    transactional_source_cols: List[str] = field(default_factory=lambda: [
        "PO Number", "Accounting Period", "AP Voucher Number",
        "Vendor Name", "WBS Element", "GL Invoice Date",
        "GL Posting Date", "GL Line Description", "Description",
        "GL Transaction Amount", "GL BER Corp Amount",
        "AP01", "AP02", "AP03", "Type"
    ])


@dataclass
class AppConfig:
    template: TemplateConfig = field(default_factory=TemplateConfig)
    forecast_reader: ForecastConfig = field(default_factory=ForecastConfig)
    transactional_detail_reader: TransactionalConfig = field(default_factory=TransactionalConfig)
    template_writer: WriterConfig = field(default_factory=WriterConfig)


CONFIG_PATH = Path("configs/config_streamlit.yaml")


def load_config() -> AppConfig:
    """
    Load config from configs/config_streamlit.yaml if it exists,
    otherwise fall back to hardcoded defaults.
    """
    if not CONFIG_PATH.exists():
        return AppConfig()
    try:
        with open(CONFIG_PATH, "r") as f:
            data = yaml.safe_load(f) or {}
        return _dict_to_config(data)
    except Exception as e:
        print(f"Warning: could not load config ({e}), using defaults")
        return AppConfig()


def config_to_dict(config: AppConfig) -> Dict[str, Any]:
    return {
        'template': asdict(config.template),
        'forecast_reader': asdict(config.forecast_reader),
        'transactional_detail_reader': asdict(config.transactional_detail_reader),
        'template_writer': asdict(config.template_writer),
    }


def _dict_to_config(data: Dict[str, Any]) -> AppConfig:
    import dataclasses

    def filter_fields(cls, raw: dict) -> dict:
        known = {f.name for f in dataclasses.fields(cls)}
        return {k: v for k, v in raw.items() if k in known}

    return AppConfig(
        template=TemplateConfig(**filter_fields(TemplateConfig, data.get('template', {}))),
        forecast_reader=ForecastConfig(**filter_fields(ForecastConfig, data.get('forecast_reader', {}))),
        transactional_detail_reader=TransactionalConfig(
            **filter_fields(TransactionalConfig, data.get('transactional_detail_reader', {}))
        ),
        template_writer=WriterConfig(**filter_fields(WriterConfig, data.get('template_writer', {}))),
    )
