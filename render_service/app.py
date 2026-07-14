import sys, os
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


if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='127.0.0.1', port=5091)  
