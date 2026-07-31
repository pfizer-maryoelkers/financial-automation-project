"""
FastAPI entry point for the Financial Automation pipeline.

Endpoints
---------
GET  /health                   liveness check
POST /extract-cost-centers     upload template → list of cost centers
POST /run                      upload all 3 files → download output .xlsx

Start the server:
    uvicorn api:app --reload --port 8000
"""

import traceback
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from api_backend import FileHandler, PipelineOrchestrator
from api_config import config_to_dict, load_config
from src.template_reader import TemplateReader

app = FastAPI(
    title="Financial Automation API",
    description="Backend pipeline for financial report generation.",
    version="1.0.0",
)


# ---------------------------------------------------------------------------
# Health check
# ---------------------------------------------------------------------------

@app.get("/health", summary="Liveness check")
def health():
    return {"status": "ok"}


# ---------------------------------------------------------------------------
# Extract cost centers from a template file
# ---------------------------------------------------------------------------

@app.post("/extract-cost-centers", summary="Upload template → return cost center list")
async def extract_cost_centers(
    template_file: UploadFile = File(..., description="Template .xlsx file"),
):
    """
    Accepts a template Excel file and returns the cost centers found in it.
    The frontend uses this to let the user pick which cost centers to process.
    """
    if not template_file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Template must be an .xlsx file")

    config = load_config()
    template_bytes = await template_file.read()

    # Write to a temp file so TemplateReader can open it
    temp_dir, file_paths = FileHandler.save_files(
        template_bytes=template_bytes,
        template_name=template_file.filename,
        forecast_files=[],           # not needed for this endpoint
        transactional_bytes=b"",
        transactional_name="placeholder.xlsx",
    )

    try:
        reader = TemplateReader(
            file_path=file_paths["template"],
            header_row=config.template.header_row,
            po_col=config.template.po_col,
            po_stop_marker=config.template.po_stop_marker,
            cost_center_col=config.template.cost_center_col,
            cost_center_start_row=config.template.cost_center_start_row,
        )
        return {"cost_centers": reader.cost_centers}
    except Exception as e:
        raise HTTPException(status_code=422, detail=str(e))
    finally:
        FileHandler.cleanup(temp_dir)


# ---------------------------------------------------------------------------
# Run the full pipeline
# ---------------------------------------------------------------------------

@app.post("/run", summary="Run pipeline → download output .xlsx")
async def run_pipeline(
    template_file: UploadFile = File(..., description="Template .xlsx file"),
    transactional_file: UploadFile = File(..., description="Transactional detail .xlsx file"),
    forecast_files: List[UploadFile] = File(..., description="One or more forecast .xlsx files"),
    selected_cost_centers: Optional[str] = Form(
        None,
        description="Comma-separated cost center IDs to process. Leave blank for all.",
    ),
):
    """
    Runs the full financial automation pipeline and returns the generated
    output file as a downloadable .xlsx attachment.

    **Form fields**
    - `template_file` — the Excel template
    - `transactional_file` — CTIES / transactional detail file
    - `forecast_files` — one or more forecast files (repeat the field for multiple)
    - `selected_cost_centers` — optional comma-separated list, e.g. `"1234,5678"`

    **Returns**  the completed output .xlsx as a file download.
    """
    # Basic validation
    for f in [template_file, transactional_file] + list(forecast_files):
        if not f.filename.endswith(".xlsx"):
            raise HTTPException(
                status_code=400,
                detail=f"All files must be .xlsx — received: {f.filename}",
            )

    # Parse optional cost center filter
    cost_centers: Optional[List[str]] = None
    if selected_cost_centers and selected_cost_centers.strip():
        cost_centers = [cc.strip() for cc in selected_cost_centers.split(",") if cc.strip()]

    # Read all uploaded bytes
    template_bytes = await template_file.read()
    trans_bytes = await transactional_file.read()
    forecast_list = [
        {"name": f.filename, "bytes": await f.read()}
        for f in forecast_files
    ]

    # Save to temp directory
    config = load_config()
    temp_dir, file_paths = FileHandler.save_files(
        template_bytes=template_bytes,
        template_name=template_file.filename,
        forecast_files=forecast_list,
        transactional_bytes=trans_bytes,
        transactional_name=transactional_file.filename,
    )

    try:
        orchestrator = PipelineOrchestrator(config=config_to_dict(config))
        output_path = orchestrator.run(
            file_paths=file_paths,
            selected_cost_centers=cost_centers,
        )

        if not Path(output_path).exists():
            raise HTTPException(status_code=500, detail="Output file was not created")

        return FileResponse(
            path=output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=Path(output_path).name,
            background=None,   # keep temp dir alive until response is sent
        )

    except HTTPException:
        FileHandler.cleanup(temp_dir)
        raise
    except Exception as e:
        FileHandler.cleanup(temp_dir)
        raise HTTPException(status_code=500, detail=f"Pipeline failed: {str(e)}\n{traceback.format_exc()}")


# ---------------------------------------------------------------------------
# Run status / exception summary (optional — useful for Build Me dashboard)
# ---------------------------------------------------------------------------

@app.post("/run/summary", summary="Run pipeline → return JSON summary instead of file")
async def run_pipeline_summary(
    template_file: UploadFile = File(...),
    transactional_file: UploadFile = File(...),
    forecast_files: List[UploadFile] = File(...),
    selected_cost_centers: Optional[str] = Form(None),
):
    """
    Same as /run but returns a JSON summary of the run (exception counts,
    execution time, log) rather than the output file.
    Useful for Build Me to display results before the user downloads the file.
    """
    cost_centers: Optional[List[str]] = None
    if selected_cost_centers and selected_cost_centers.strip():
        cost_centers = [cc.strip() for cc in selected_cost_centers.split(",") if cc.strip()]

    template_bytes = await template_file.read()
    trans_bytes = await transactional_file.read()
    forecast_list = [
        {"name": f.filename, "bytes": await f.read()}
        for f in forecast_files
    ]

    config = load_config()
    temp_dir, file_paths = FileHandler.save_files(
        template_bytes=template_bytes,
        template_name=template_file.filename,
        forecast_files=forecast_list,
        transactional_bytes=trans_bytes,
        transactional_name=transactional_file.filename,
    )

    try:
        orchestrator = PipelineOrchestrator(config=config_to_dict(config))
        output_path = orchestrator.run(
            file_paths=file_paths,
            selected_cost_centers=cost_centers,
        )

        output_bytes = Path(output_path).read_bytes()

        return JSONResponse({
            "status": "success",
            "output_filename": Path(output_path).name,
            "execution_time_seconds": orchestrator.logger.execution_time,
            "exceptions": orchestrator.get_exception_summary(),
            "logs": orchestrator.logger.logs,
            # base64-encode the file so Build Me can offer a download link
            "output_file_base64": __import__("base64").b64encode(output_bytes).decode(),
        })

    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "status": "error",
                "detail": str(e),
                "traceback": traceback.format_exc(),
            },
        )
    finally:
        FileHandler.cleanup(temp_dir)
