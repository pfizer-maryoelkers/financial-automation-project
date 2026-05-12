"""
Configuration management for Streamlit UI.
Handles loading, saving, validation, and defaults for all pipeline configurations.
"""

from dataclasses import dataclass, field, asdict
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path
import yaml
import re


@dataclass
class TemplateConfig:
    """Template-related configuration"""
    header_row: int = 16
    po_col: str = "B"
    po_stop_marker: str = "Previous Period Invoices"
    cost_center_col: str = "A"
    cost_center_start_row: int = 9


@dataclass
class ForecastConfig:
    """Forecast reader configuration"""
    po_col: str = "PO #"


@dataclass
class TransactionalConfig:
    """Transactional detail reader configuration"""
    required_cols: List[str] = field(default_factory=lambda: [
        "PO Number", "Month", "GL Transaction Amount"
    ])
    valid_types: List[str] = field(default_factory=lambda: [
        "Actual", "Accrual", "Reversal"
    ])
    colmap: Dict[str, str] = field(default_factory=lambda: {
        'po': 'PO Number',
        'month': 'Month',
        'amount': 'GL BER Corp Amount',
        'classifier': 'AP Voucher Number',
        'cost_center': 'Cost Center*',
        'wbs': 'WBS Element',
        'type': 'Type'
    })


@dataclass
class WriterConfig:
    """Template writer configuration"""
    output_path: str = "template_output.xlsx"
    overwrite: bool = False
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
        "GL Transaction Amount", "GL BER Corp Amount", "Month",
        "AP01", "AP02", "AP03", "Type"
    ])


@dataclass
class AppConfig:
    """Complete application configuration"""
    template: TemplateConfig = field(default_factory=TemplateConfig)
    forecast_reader: ForecastConfig = field(default_factory=ForecastConfig)
    transactional_detail_reader: TransactionalConfig = field(default_factory=TransactionalConfig)
    template_writer: WriterConfig = field(default_factory=WriterConfig)


class ConfigManager:
    """Manages configuration loading, saving, and validation"""
    
    STREAMLIT_CONFIG_PATH = Path("configs/config_streamlit.yaml")
    
    @staticmethod
    def get_default_config() -> AppConfig:
        """Get default configuration"""
        return AppConfig()
    
    @staticmethod
    def load_config() -> AppConfig:
        """
        Load configuration with priority:
        1. config_streamlit.yaml (if exists)
        2. Default hardcoded values
        """
        if ConfigManager.STREAMLIT_CONFIG_PATH.exists():
            try:
                with open(ConfigManager.STREAMLIT_CONFIG_PATH, 'r') as f:
                    data = yaml.safe_load(f)
                return ConfigManager._dict_to_config(data)
            except Exception as e:
                print(f"Error loading config: {e}, using defaults")
                return ConfigManager.get_default_config()
        return ConfigManager.get_default_config()
    
    @staticmethod
    def save_config(config: AppConfig) -> bool:
        """Save configuration to config_streamlit.yaml"""
        try:
            ConfigManager.STREAMLIT_CONFIG_PATH.parent.mkdir(exist_ok=True)
            data = ConfigManager._config_to_dict(config)
            with open(ConfigManager.STREAMLIT_CONFIG_PATH, 'w') as f:
                yaml.dump(data, f, default_flow_style=False, sort_keys=False)
            return True
        except Exception as e:
            print(f"Error saving config: {e}")
            return False
    
    @staticmethod
    def validate_config(config: AppConfig) -> Tuple[bool, List[str]]:
        """
        Validate configuration values.
        Returns (is_valid, list_of_errors)
        """
        errors = []
        
        # Template validation
        if config.template.header_row < 1:
            errors.append("Template header row must be >= 1")
        if not ConfigManager._is_valid_excel_column(config.template.po_col):
            errors.append(f"Invalid PO column: {config.template.po_col}")
        if not ConfigManager._is_valid_excel_column(config.template.cost_center_col):
            errors.append(f"Invalid cost center column: {config.template.cost_center_col}")
        if config.template.cost_center_start_row < 1:
            errors.append("Cost center start row must be >= 1")
        if not config.template.po_stop_marker.strip():
            errors.append("PO stop marker cannot be empty")
        
        # Forecast validation
        if not config.forecast_reader.po_col.strip():
            errors.append("Forecast PO column name cannot be empty")
        
        # Transactional validation
        if not config.transactional_detail_reader.required_cols:
            errors.append("At least one required column must be specified")
        if not config.transactional_detail_reader.valid_types:
            errors.append("At least one valid type must be specified")
        for key, value in config.transactional_detail_reader.colmap.items():
            if not value.strip():
                errors.append(f"Transactional column mapping '{key}' cannot be empty")
        
        # Writer validation
        if not config.template_writer.output_path.endswith('.xlsx'):
            errors.append("Output filename must end with .xlsx")
        if not ConfigManager._is_valid_excel_column(config.template_writer.dec_acc_reversal_col):
            errors.append(f"Invalid Dec Acc Reversal column: {config.template_writer.dec_acc_reversal_col}")
        if not config.template_writer.forecast_source_cols:
            errors.append("At least one forecast source column must be specified")
        if not config.template_writer.transactional_source_cols:
            errors.append("At least one transactional source column must be specified")
        
        return len(errors) == 0, errors
    
    @staticmethod
    def _is_valid_excel_column(col: str) -> bool:
        """Validate Excel column format (A-ZZ)"""
        return bool(re.match(r'^[A-Z]{1,2}$', col.upper()))
    
    @staticmethod
    def _config_to_dict(config: AppConfig) -> Dict[str, Any]:
        """Convert AppConfig to dictionary for YAML"""
        return {
            'template': asdict(config.template),
            'forecast_reader': asdict(config.forecast_reader),
            'transactional_detail_reader': asdict(config.transactional_detail_reader),
            'template_writer': asdict(config.template_writer)
        }
    
    @staticmethod
    def _dict_to_config(data: Dict[str, Any]) -> AppConfig:
        """Convert dictionary to AppConfig"""
        return AppConfig(
            template=TemplateConfig(**data.get('template', {})),
            forecast_reader=ForecastConfig(**data.get('forecast_reader', {})),
            transactional_detail_reader=TransactionalConfig(**data.get('transactional_detail_reader', {})),
            template_writer=WriterConfig(**data.get('template_writer', {}))
        )
    
    @staticmethod
    def export_config_yaml(config: AppConfig) -> str:
        """Export configuration as YAML string for download"""
        data = ConfigManager._config_to_dict(config)
        return yaml.dump(data, default_flow_style=False, sort_keys=False)
    
    @staticmethod
    def is_using_custom_config() -> bool:
        """Check if custom config file exists"""
        return ConfigManager.STREAMLIT_CONFIG_PATH.exists()
    
    @staticmethod
    def delete_custom_config() -> bool:
        """Delete custom config file"""
        try:
            if ConfigManager.STREAMLIT_CONFIG_PATH.exists():
                ConfigManager.STREAMLIT_CONFIG_PATH.unlink()
            return True
        except Exception as e:
            print(f"Error deleting config: {e}")
            return False

# Made with Bob
