"""
Pipeline orchestrator for the FastAPI layer.
Mirrors streamlit_backend.py but has no Streamlit dependency.
All progress is reported via a simple callback float (0.0–1.0).
"""

import shutil
import tempfile
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional

from src.forecast_reader import ForecastReader
from src.models import ExceptionLog
from src.template_reader import TemplateReader
from src.template_writer import TemplateWriter
from src.transactional_detail_reader import TransactionalDetailReader
from src.utils import build_hierarchy


# ---------------------------------------------------------------------------
# Simple logger — collects structured log entries, no UI dependency
# ---------------------------------------------------------------------------

class PipelineLogger:
    def __init__(self):
        self.logs: List[Dict] = []
        self._start: Optional[float] = None

    def start(self):
        self._start = time.time()
        self.logs = []

    def _append(self, level: str, message: str):
        self.logs.append({
            "level": level,
            "timestamp": datetime.now().strftime("%H:%M:%S"),
            "message": message,
        })
        print(f"[{level}] {message}")

    def info(self, msg: str):    self._append("INFO", msg)
    def success(self, msg: str): self._append("SUCCESS", msg)
    def warning(self, msg: str): self._append("WARNING", msg)
    def error(self, msg: str):   self._append("ERROR", msg)

    @property
    def execution_time(self) -> Optional[float]:
        return round(time.time() - self._start, 2) if self._start else None


# ---------------------------------------------------------------------------
# File helpers
# ---------------------------------------------------------------------------

class FileHandler:
    """Save uploaded raw bytes to a temp directory; clean up afterwards."""

    @staticmethod
    def save_files(
        template_bytes: bytes,
        template_name: str,
        forecast_files: List[Dict],   # [{"name": str, "bytes": bytes}, ...]
        transactional_bytes: bytes,
        transactional_name: str,
    ) -> tuple[str, Dict]:
        """
        Write uploaded file bytes to a temp directory.

        Returns
        -------
        temp_dir : str
        file_paths : {"template": str, "forecast": [str, ...], "transactional": str}
        """
        temp_dir = tempfile.mkdtemp(prefix="finance_api_")
        try:
            template_path = Path(temp_dir) / template_name
            template_path.write_bytes(template_bytes)

            forecast_paths = []
            for i, f in enumerate(forecast_files):
                p = Path(temp_dir) / f"forecast_{i}_{f['name']}"
                p.write_bytes(f["bytes"])
                forecast_paths.append(str(p))

            trans_path = Path(temp_dir) / transactional_name
            trans_path.write_bytes(transactional_bytes)

            return temp_dir, {
                "template": str(template_path),
                "forecast": forecast_paths,
                "transactional": str(trans_path),
            }
        except Exception:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise

    @staticmethod
    def cleanup(temp_dir: str):
        shutil.rmtree(temp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Pipeline orchestrator
# ---------------------------------------------------------------------------

class PipelineOrchestrator:
    """
    Runs the full financial automation pipeline.
    Drop-in replacement for the Streamlit orchestrator — no UI imports.
    """

    def __init__(
        self,
        config: Dict,
        progress_callback: Optional[Callable[[int], None]] = None,
    ):
        self.config = config
        self.progress_callback = progress_callback
        self.logger = PipelineLogger()
        self.exception_log: Optional[ExceptionLog] = None
        self.output_path: Optional[str] = None

    def _progress(self, pct: int):
        if self.progress_callback:
            self.progress_callback(pct)

    def run(
        self,
        file_paths: Dict,
        selected_cost_centers: Optional[List[str]] = None,
    ) -> str:
        """
        Execute all pipeline steps.

        Parameters
        ----------
        file_paths : {"template": str, "forecast": [str,...], "transactional": str}
        selected_cost_centers : optional filter list

        Returns
        -------
        str  path to the generated output .xlsx file
        """
        self.logger.start()
        self.logger.info("Pipeline starting...")
        self._progress(0)

        # -- Step 1: Load data -----------------------------------------------
        self.logger.info("Step 1/4: Loading data...")
        self._progress(10)

        forecast_reader = ForecastReader(
            file_paths=(
                file_paths["forecast"]
                if isinstance(file_paths["forecast"], list)
                else [file_paths["forecast"]]
            ),
            po_col=self.config["forecast_reader"]["po_col"],
        )
        forecast_data = forecast_reader.get_forecast_data()
        self.logger.info(f"Forecast loaded: {len(forecast_data)} POs")
        self._progress(20)

        transactional_reader = TransactionalDetailReader(
            file_path=file_paths["transactional"],
            required_cols=self.config["transactional_detail_reader"]["required_cols"],
            valid_types=self.config["transactional_detail_reader"]["valid_types"],
            colmap=self.config["transactional_detail_reader"]["colmap"],
        )
        transactional_data = transactional_reader.get_transactional_data()
        reclass_data = transactional_reader.get_reclass_data()
        reclass_notes = transactional_reader.get_reclass_notes()
        hierarchy_map = transactional_reader.get_hierarchy_map()
        intl_po_set = transactional_reader.get_intl_po_set()
        row_count = len(transactional_reader.data) if transactional_reader.data is not None else 0
        self.logger.info(f"Transactional loaded: {row_count} rows")
        self._progress(30)

        template_reader = TemplateReader(
            file_path=file_paths["template"],
            header_row=self.config["template"]["header_row"],
            po_col=self.config["template"]["po_col"],
            po_stop_marker=self.config["template"]["po_stop_marker"],
            cost_center_col=self.config["template"]["cost_center_col"],
            cost_center_start_row=self.config["template"]["cost_center_start_row"],
        )
        self.logger.info(f"Template loaded: {len(template_reader.cost_centers)} cost centers")
        self._progress(40)

        cost_centers_to_process = template_reader.cost_centers
        if selected_cost_centers:
            cost_centers_to_process = [
                cc for cc in template_reader.cost_centers if cc in selected_cost_centers
            ]
            self.logger.info(f"Filtered to {len(cost_centers_to_process)} cost centers")

        # -- Step 2: Build hierarchy -----------------------------------------
        self.logger.info("Step 2/4: Building hierarchy...")
        self._progress(45)
        self.exception_log = ExceptionLog()

        if transactional_reader.data is None:
            raise RuntimeError("Transactional data failed to load")

        hierarchy = build_hierarchy(
            cost_centers=cost_centers_to_process,
            hierarchy_map=hierarchy_map,
            transactional_data=transactional_data,
            forecast_data=forecast_data,
            exception_log=self.exception_log,
            transactional_df=transactional_reader.data,
            reclass_data=reclass_data,
            reclass_notes=reclass_notes,
            template_pos=template_reader.pos,
            template_rows=template_reader.template_rows,
            intl_po_set=intl_po_set,
        )
        self.logger.info(f"Hierarchy built: {len(self.exception_log.entries)} exceptions")
        self._progress(60)

        # -- Step 3: Write template ------------------------------------------
        self.logger.info("Step 3/4: Writing output...")
        self._progress(65)

        output_filename = Path(
            self.config["template_writer"].get("output_path", "template_output.xlsx")
        ).name
        output_path = Path(file_paths["template"]).parent / output_filename

        template_writer = TemplateWriter(
            file_path=file_paths["template"],
            header_row=self.config["template"]["header_row"],
            po_column=self.config["template"]["po_col"],
            output_path=str(output_path),
            overwrite=self.config["template_writer"]["overwrite"],
            dec_acc_reversal_col=self.config["template_writer"]["dec_acc_reversal_col"],
            forecast_source_cols=self.config["template_writer"]["forecast_source_cols"],
            transactional_source_cols=self.config["template_writer"]["transactional_source_cols"],
        )
        self._progress(70)

        pos = template_writer.insert_missing_po_rows(hierarchy, pos=template_reader.pos)
        pos = template_writer.insert_er_rows(hierarchy, pos=pos)
        template_writer.write_hierarchy(hierarchy, pos=pos)
        self._progress(80)
        self.logger.info("Template written")
        self._progress(85)

        # -- Step 4: Exception reports ----------------------------------------
        self.logger.info("Step 4/4: Writing exception reports...")
        self._progress(90)
        template_writer.write_exception_data_sheet(self.exception_log)
        template_writer.write_exception_sheet(self.exception_log, transactional_reader.data)
        self._progress(95)
        template_writer.save()
        self._progress(100)

        self.output_path = str(output_path)
        self.logger.success(
            f"Pipeline complete in {self.logger.execution_time}s — "
            f"{len(self.exception_log.entries)} exceptions"
        )
        return self.output_path

    def get_exception_summary(self) -> Optional[Dict]:
        return self.exception_log.summary_by_type() if self.exception_log else None
