"""
Backend orchestrator for Streamlit UI.
Handles file management, pipeline execution, and logging.
"""

import tempfile
import shutil
import time
from pathlib import Path
from typing import Optional, Dict, List, Tuple
from datetime import datetime

from src.forecast_reader import ForecastReader
from src.transactional_detail_reader import TransactionalDetailReader
from src.template_reader import TemplateReader
from src.template_writer import TemplateWriter
from src.utils import build_hierarchy
from src.models import ExceptionLog


class FileHandler:
    """Handles temporary file storage and cleanup for uploaded files."""
    
    @staticmethod
    def save_uploaded_files(
        template_file,
        forecast_files: List,
        transactional_file
    ) -> Tuple[str, Dict[str, str]]:
        """
        Save uploaded files to a temporary directory.
        
        Args:
            template_file: Uploaded template file
            forecast_files: List of uploaded forecast files
            transactional_file: Uploaded transactional detail file
            
        Returns:
            Tuple of (temp_dir_path, file_paths_dict)
        """
        temp_dir = tempfile.mkdtemp(prefix='streamlit_finance_')
        file_paths = {}
        
        try:
            # Save template file
            if template_file is not None:
                template_path = Path(temp_dir) / template_file.name
                with open(template_path, 'wb') as f:
                    f.write(template_file.getbuffer())
                file_paths['template'] = str(template_path)
            
            # Save forecast files
            if forecast_files:
                forecast_paths = []
                for i, forecast_file in enumerate(forecast_files):
                    forecast_path = Path(temp_dir) / f"forecast_{i}_{forecast_file.name}"
                    with open(forecast_path, 'wb') as f:
                        f.write(forecast_file.getbuffer())
                    forecast_paths.append(str(forecast_path))
                file_paths['forecast'] = forecast_paths
            
            # Save transactional file
            if transactional_file is not None:
                trans_path = Path(temp_dir) / transactional_file.name
                with open(trans_path, 'wb') as f:
                    f.write(transactional_file.getbuffer())
                file_paths['transactional'] = str(trans_path)
            
            return temp_dir, file_paths
            
        except Exception as e:
            # Clean up on error
            FileHandler.cleanup_temp_files(temp_dir)
            raise Exception(f"Error saving uploaded files: {str(e)}")
    
    @staticmethod
    def cleanup_temp_files(temp_dir: str) -> None:
        """
        Remove temporary directory and all its contents.
        
        Args:
            temp_dir: Path to temporary directory
        """
        if temp_dir and Path(temp_dir).exists():
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                print(f"Warning: Could not clean up temp directory {temp_dir}: {str(e)}")
    
    @staticmethod
    def get_output_file_bytes(output_path: str) -> Optional[bytes]:
        """
        Read output file as bytes for download.
        
        Args:
            output_path: Path to output file
            
        Returns:
            File contents as bytes, or None if file doesn't exist
        """
        try:
            if Path(output_path).exists():
                with open(output_path, 'rb') as f:
                    return f.read()
            return None
        except Exception as e:
            print(f"Error reading output file: {str(e)}")
            return None


class StreamlitLogger:
    """Custom logger for Streamlit UI with progress tracking."""
    
    def __init__(self, status_container=None, log_container=None):
        """
        Initialize logger.
        
        Args:
            status_container: Streamlit container for status messages
            log_container: Streamlit container for detailed logs
        """
        self.status_container = status_container
        self.log_container = log_container
        self.logs = []
        self.start_time = None
    
    def start(self):
        """Mark the start of pipeline execution."""
        self.start_time = time.time()
        self.logs = []
    
    def info(self, message: str):
        """Log an info message."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.logs.append(('INFO', timestamp, message))
        if self.status_container:
            self.status_container.info(message)
    
    def success(self, message: str):
        """Log a success message."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.logs.append(('SUCCESS', timestamp, message))
        if self.status_container:
            self.status_container.success(message)
    
    def error(self, message: str):
        """Log an error message."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.logs.append(('ERROR', timestamp, message))
        if self.status_container:
            self.status_container.error(message)
    
    def warning(self, message: str):
        """Log a warning message."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.logs.append(('WARNING', timestamp, message))
        if self.status_container:
            self.status_container.warning(message)
    
    def get_logs(self) -> List[Tuple[str, str, str]]:
        """Get all logged messages."""
        return self.logs
    
    def get_execution_time(self) -> Optional[float]:
        """Get execution time in seconds."""
        if self.start_time:
            return time.time() - self.start_time
        return None


class PipelineOrchestrator:
    """Orchestrates the execution of the financial automation pipeline."""
    
    def __init__(self, config: Dict, logger: Optional[StreamlitLogger] = None, progress_callback=None):
        """
        Initialize orchestrator.
        
        Args:
            config: Configuration dictionary
            logger: Optional StreamlitLogger instance
            progress_callback: Optional callback function for progress updates (receives percentage 0-100)
        """
        self.config = config
        self.logger = logger or StreamlitLogger()
        self.progress_callback = progress_callback
        self.exception_log = None
        self.output_path = None
    
    def _update_progress(self, percentage: int):
        """Update progress if callback is provided."""
        if self.progress_callback:
            self.progress_callback(percentage)
    
    def run(self, file_paths: Dict[str, str], selected_cost_centers: Optional[List[str]] = None) -> str:
        """
        Execute the pipeline with uploaded files.
        
        Args:
            file_paths: Dictionary with keys 'template', 'forecast', 'transactional'
            selected_cost_centers: Optional list of cost centers to process. If None, processes all.
            
        Returns:
            Path to generated output file
            
        Raises:
            Exception: If pipeline execution fails
        """
        try:
            self.logger.start()
            self.logger.info("Starting pipeline execution...")
            self._update_progress(0)
            
            # Step 1: Initialize readers and load data
            self.logger.info("Step 1/4: Loading data...")
            self._update_progress(10)
            
            forecast_reader = ForecastReader(
                file_paths=file_paths['forecast'] if isinstance(file_paths['forecast'], list) else [file_paths['forecast']],
                po_col=self.config['forecast_reader']['po_col']
            )
            forecast_data = forecast_reader.get_forecast_data()
            self.logger.info(f"Loaded forecast data: {len(forecast_data)} POs")
            self._update_progress(20)
            
            transactional_reader = TransactionalDetailReader(
                file_path=file_paths['transactional'],
                required_cols=self.config['transactional_detail_reader']['required_cols'],
                valid_types=self.config['transactional_detail_reader']['valid_types'],
                colmap=self.config['transactional_detail_reader']['colmap']
            )
            transactional_data = transactional_reader.get_transactional_data()
            hierarchy_map = transactional_reader.get_hierarchy_map()
            row_count = len(transactional_reader.data) if transactional_reader.data is not None else 0
            self.logger.info(f"Loaded transactional data: {row_count} rows")
            self._update_progress(30)
            
            template_reader = TemplateReader(
                file_path=file_paths['template'],
                header_row=self.config['template']['header_row'],
                po_col=self.config['template']['po_col'],
                po_stop_marker=self.config['template']['po_stop_marker'],
                cost_center_col=self.config['template']['cost_center_col'],
                cost_center_start_row=self.config['template']['cost_center_start_row']
            )
            self.logger.info(f"Loaded template: {len(template_reader.cost_centers)} cost centers")
            self._update_progress(40)
            
            # Filter cost centers if selection provided
            cost_centers_to_process = template_reader.cost_centers
            if selected_cost_centers:
                cost_centers_to_process = [cc for cc in template_reader.cost_centers if cc in selected_cost_centers]
                self.logger.info(f"Filtering to {len(cost_centers_to_process)} selected cost centers")
            
            # Step 2: Build hierarchy
            self.logger.info("Step 2/4: Building hierarchy...")
            self._update_progress(45)
            self.exception_log = ExceptionLog()
            
            # Ensure transactional data is loaded
            if transactional_reader.data is None:
                raise Exception("Transactional data failed to load")
            
            hierarchy = build_hierarchy(
                cost_centers=cost_centers_to_process,
                hierarchy_map=hierarchy_map,
                transactional_data=transactional_data,
                forecast_data=forecast_data,
                exception_log=self.exception_log,
                transactional_df=transactional_reader.data
            )
            
            total_exceptions = len(self.exception_log.entries)
            self.logger.info(f"Built hierarchy: {total_exceptions} exceptions found")
            self._update_progress(60)
            
            # Step 3: Write to template
            self.logger.info("Step 3/4: Writing template output...")
            self._update_progress(65)
            
            # Determine output path - use just the filename in temp directory
            output_filename = self.config['template_writer'].get('output_path', 'template_output.xlsx')
            # Extract just the filename if it's a path
            output_filename = Path(output_filename).name
            output_path = Path(file_paths['template']).parent / output_filename
            
            template_writer = TemplateWriter(
                file_path=file_paths['template'],
                header_row=self.config['template']['header_row'],
                po_column=self.config['template']['po_col'],
                output_path=str(output_path),
                overwrite=self.config['template_writer']['overwrite'],
                dec_acc_reversal_col=self.config['template_writer']['dec_acc_reversal_col'],
                forecast_source_cols=self.config['template_writer']['forecast_source_cols'],
                transactional_source_cols=self.config['template_writer']['transactional_source_cols']
            )
            self._update_progress(70)
            
            template_writer.write_hierarchy(hierarchy, pos=template_reader.pos)
            self._update_progress(75)
            template_writer.write_forecast_source_sheet(forecast_reader.data, pos=template_reader.pos)
            self._update_progress(80)
            template_writer.write_transactional_source_sheet(transactional_reader.data, pos=template_reader.pos)
            self.logger.info("Template data written")
            self._update_progress(85)
            
            # Step 4: Generate exception reports
            self.logger.info("Step 4/4: Generating exception reports...")
            self._update_progress(90)
            
            template_writer.write_exception_data_sheet(self.exception_log)
            template_writer.write_exception_sheet(self.exception_log, transactional_reader.data)
            template_writer.write_exception_summary_sheet(self.exception_log)
            self._update_progress(95)
            template_writer.save()
            
            self.logger.info("Exception reports generated")
            self._update_progress(100)
            
            # Success
            self.output_path = str(output_path)
            execution_time = self.logger.get_execution_time()
            self.logger.success(f"Pipeline completed successfully in {execution_time:.2f} seconds!")
            
            return self.output_path
            
        except Exception as e:
            self.logger.error(f"Pipeline failed: {str(e)}")
            raise
    
    def get_exception_summary(self) -> Optional[Dict]:
        """
        Get summary of exceptions found during pipeline execution.
        
        Returns:
            Dictionary with exception counts by type, or None if not available
        """
        if self.exception_log:
            return self.exception_log.summary_by_type()
        return None

# Made with Bob


class ExcelPreviewHandler:
    """Handles Excel file preview functionality."""
    
    @staticmethod
    def get_sheet_names(file_path: str) -> List[str]:
        """
        Get all sheet names from Excel file.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            List of sheet names
        """
        try:
            import pandas as pd
            with pd.ExcelFile(file_path) as xls:
                return xls.sheet_names
        except Exception as e:
            print(f"Error reading sheet names: {e}")
            return []
    
    @staticmethod
    def get_sheet_info(file_path: str, sheet_name: str) -> Dict[str, int]:
        """
        Get metadata about a specific sheet.
        
        Args:
            file_path: Path to Excel file
            sheet_name: Name of sheet to analyze
            
        Returns:
            Dictionary with row_count and column_count
        """
        try:
            import pandas as pd
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            return {
                'row_count': len(df),
                'column_count': len(df.columns)
            }
        except Exception as e:
            print(f"Error reading sheet info: {e}")
            return {'row_count': 0, 'column_count': 0}
    
    @staticmethod
    def preview_sheet(file_path: str, sheet_name: str, max_rows: int = 50, header_row: Optional[int] = None):
        """
        Load first N rows of a sheet for preview.
        
        Args:
            file_path: Path to Excel file
            sheet_name: Name of sheet to preview
            max_rows: Maximum number of rows to load
            header_row: Optional header row number (0-based). If provided, uses this row as column headers.
            
        Returns:
            DataFrame with preview data, or None if error
        """
        try:
            import pandas as pd
            # If header_row is provided, convert from 1-based to 0-based and use it
            if header_row is not None:
                header_idx = header_row - 1  # Convert from 1-based to 0-based
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    header=header_idx,
                    nrows=max_rows
                )
            else:
                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    nrows=max_rows
                )
            return df
        except Exception as e:
            print(f"Error loading sheet preview: {e}")
            return None
