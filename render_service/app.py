import sys, os, traceback
import tempfile
import shutil
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(_THIS_DIR.parent))

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from render_service.models import RenderRequest, RenderResponse
from render_service.renderer import render, render_with_fields
from render_service.excel_cluster_reader import read_fields
from render_service.upload_coordinate_generator import generate_coordinates, generate_coordinates_and_preview

app = FastAPI(title='PaperLess Render Service', version='1.0.0')

# CORS: allow the Next.js frontend during development
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:3002",
    ],
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


@app.get('/health')
def health():
    return {'status': 'ok'}


@app.post('/render/runtime', response_model=RenderResponse)
def render_runtime(req: RenderRequest):
    try:
        return render(template_id=req.template_id, xlsx_path=req.xlsx_path, output_dir=req.output_dir)
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post('/upload', response_model=RenderResponse)
async def upload_runtime(file: UploadFile = File(...)):
    """
    Upload an Excel workbook and render it with fields read from the
    hidden _Fields sheet (or cell comments as fallback).

    No database lookup is performed — field definitions come entirely
    from the uploaded file.
    """
    tmp_dir = tempfile.mkdtemp(prefix="ple_upload_")
    try:
        # Save uploaded file to temp location
        xlsx_path = os.path.join(tmp_dir, file.filename or "upload.xlsx")
        content = await file.read()
        with open(xlsx_path, "wb") as f:
            f.write(content)

        # Read cluster metadata from the workbook
        print(f"[upload] Reading cluster metadata from: {xlsx_path}")
        try:
            fields = read_fields(xlsx_path)
        except Exception as e:
            error_msg = str(e)
            print(f"[upload] Cluster read failed: {error_msg}")

            # Distinguish COM init failure (infrastructure) from invalid workbook (user error)
            if "CoInitialize" in error_msg:
                raise HTTPException(
                    status_code=500,
                    detail={
                        "success": False,
                        "message": "Excel COM initialization failed.",
                        "detail": error_msg,
                    },
                )
            elif (
                "cannot open" in error_msg.lower()
                or "not a valid" in error_msg.lower()
                or "file not found" in error_msg.lower()
                or "cannot access" in error_msg.lower()
            ):
                raise HTTPException(
                    status_code=400,
                    detail={
                        "success": False,
                        "message": "Unable to open workbook.",
                    },
                )
            else:
                # Assume workbook opened but metadata extraction failed
                raise HTTPException(
                    status_code=400,
                    detail={
                        "success": False,
                        "message": (
                            "This workbook does not contain ConMas field metadata "
                            "(_Fields sheet or supported metadata source)."
                        ),
                    },
                )

        if not fields:
            raise HTTPException(
                status_code=400,
                detail={
                    "success": False,
                    "message": (
                        "This workbook does not contain ConMas field metadata "
                        "(_Fields sheet or supported metadata source)."
                    ),
                },
            )
        print(f"[upload] Found {len(fields)} fields from workbook metadata")

        # Render using the upload path (no DB, no calibration)
        try:
            result = render_with_fields(
                fields=fields,
                xlsx_path=xlsx_path,
                dpi=300,
            )
        except Exception as e:
            print(f"[upload] Render failed: {e}")
            raise HTTPException(status_code=500, detail=f"Render failed: {str(e)}")

        print(f"[upload] Render complete: page={result.page_width}x{result.page_height}, "
              f"fields={len(result.fields)}")
        return result

    except HTTPException:
        raise
    except Exception as e:
        print(f"[upload] Unexpected error: {e}")
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.post('/upload/coordinates')
async def upload_coordinates(file: UploadFile = File(...)):
    """
    Generate ConMas-compatible field coordinates for a new workbook.

    Implements the exact MakeCluster → ExportPdf → GetAddress → normalize
    pipeline from the original ConMas Designer.

    Returns field definitions with normalized ratios (left_ratio, top_ratio,
    right_ratio, bottom_ratio) ready for database storage.

    No workbook geometry, column widths, or calibration is used — only
    pixel scanning of a sanitized PDF rendered at 300 DPI.
    """
    tmp_dir = tempfile.mkdtemp(prefix="ple_upload_coords_")
    try:
        xlsx_path = os.path.join(tmp_dir, file.filename or "upload.xlsx")
        content = await file.read()
        with open(xlsx_path, "wb") as f:
            f.write(content)

        print(f"[upload/coordinates] Generating coordinates for: {file.filename}")
        fields = generate_coordinates(xlsx_path)

        if not fields:
            raise HTTPException(
                status_code=400,
                detail={
                    "success": False,
                    "message": (
                        "No ConMas field comments found in this workbook. "
                        "Ensure cells have ConMas-style comments attached."
                    ),
                },
            )

        print(f"[upload/coordinates] Found {len(fields)} fields")
        return {"success": True, "fields": fields}

    except HTTPException:
        raise
    except Exception as e:
        print(f"[upload/coordinates] Error: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Coordinate generation failed: {str(e)}")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.post('/upload/preview')
async def upload_preview(
    file: UploadFile = File(...),
    output_dir: str | None = None,
):
    """
    Upload an Excel workbook and generate a preview with both background PNG
    and ConMas-compatible field coordinates.

    This is a PREVIEW-ONLY endpoint — nothing is saved to the database.

    Pipeline:
      1. Generate coordinates via MakeCluster → ExportPdf → pixel scan (ConMas)
      2. Render original workbook as PDF → 300 DPI PNG for background
      3. Return PNG filename + page dimensions + field ratios

    Args:
        file: The uploaded .xlsx file.
        output_dir: Optional directory to save the background PNG. If provided,
            the PNG is saved here and this dir must persist after the request.
            If omitted, a temp dir is used (caller must read the PNG before
            the dir is cleaned up).
    """
    tmp_dir = tempfile.mkdtemp(prefix="ple_preview_")
    try:
        xlsx_path = os.path.join(tmp_dir, file.filename or "upload.xlsx")
        content = await file.read()
        with open(xlsx_path, "wb") as f:
            f.write(content)

        out_dir = output_dir or tmp_dir
        output_id = Path(file.filename or "upload").stem
        # Single COM session: matches original ConMas architecture
        result = generate_coordinates_and_preview(
            xlsx_path, output_dir=out_dir, output_id=output_id
        )

        if not result.get("fields"):
            print(f"[upload/preview] No fields detected in: {file.filename}")

        print(f"[upload/preview] Complete: page={result['page']['width']}x{result['page']['height']}, "
              f"fields={len(result.get('fields', []))}, bg={result.get('backgroundImage')}")

        return {
            "success": True,
            "backgroundImage": result["backgroundImage"],
            "page": result["page"],
            "fields": result["fields"],
        }

    except Exception as e:
        print(f"[upload/preview] Error: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Preview generation failed: {str(e)}")
    finally:
        # Clean up temp XLSX dir when output_dir is managed by caller
        if output_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='127.0.0.1', port=5091)  
