import os
import io
import sys
import webbrowser
import uvicorn
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query
from fastapi.responses import StreamingResponse, FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from typing import Optional
import json

from processor import (
    parse_input_workbook,
    generate_sf9_front_workbook,
    generate_sf9_back_workbook,
    generate_sf9_zip_bundle,
    get_output_filenames,
    FRONT_TEMPLATE_PATH,
    BACK_TEMPLATE_PATH,
)

app = FastAPI(title="SF9 Automated Report Card Generator", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "X-Suggested-Filename", "X-Student-Count"]
)

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
INPUT_TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "SF9_Data_Input_Template.xlsx")

os.makedirs(STATIC_DIR, exist_ok=True)

@app.post("/api/preview")
async def preview_file(file: UploadFile = File(...)):
    """
    Parses the uploaded Excel file and returns student preview data, summary stats,
    and output filenames for FRONT and BACK workbooks.
    """
    if not file.filename.endswith(('.xlsx', '.xlsm', '.xltx')):
        raise HTTPException(status_code=400, detail="Please upload a valid Excel file (.xlsx).")

    contents = await file.read()
    try:
        students = parse_input_workbook(contents)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to parse Excel file: {str(e)}")

    if not students:
        return {
            "success": False,
            "message": "No valid student records found. Ensure your sheet has student names or LRNs.",
            "count": 0,
            "students": [],
            "front_filename": "SF9_FRONT_Empty.xlsx",
            "back_filename": "SF9_BACK_Empty.xlsx",
            "zip_filename": "SF9_Complete_Empty.zip"
        }

    front_name, back_name, zip_name = get_output_filenames(students)

    sections = list(set([str(s.get('section')) for s in students if s.get('section')]))
    grades = list(set([str(s.get('grade')) for s in students if s.get('grade')]))
    school_years = list(set([str(s.get('school_year')) for s in students if s.get('school_year')]))

    return {
        "success": True,
        "count": len(students),
        "front_filename": front_name,
        "back_filename": back_name,
        "zip_filename": zip_name,
        "summary": {
            "grades": grades,
            "sections": sections,
            "school_year": school_years[0] if school_years else ""
        },
        "students": students
    }

@app.post("/api/convert/front")
async def convert_front(
    file: UploadFile = File(...),
    overrides: Optional[str] = Form(None)
):
    """Generates and streams the FRONT SF9 workbook (1 sheet per student)."""
    if not file.filename.endswith(('.xlsx', '.xlsm', '.xltx')):
        raise HTTPException(status_code=400, detail="Please upload a valid Excel file (.xlsx).")

    contents = await file.read()
    students = parse_input_workbook(contents)
    if not students:
        raise HTTPException(status_code=400, detail="No student records found.")

    global_overrides = {}
    if overrides:
        try:
            global_overrides = json.loads(overrides)
        except Exception:
            pass

    try:
        wb_front, out_filename = generate_sf9_front_workbook(students, global_overrides=global_overrides)
        output_stream = io.BytesIO()
        wb_front.save(output_stream)
        output_stream.seek(0)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate FRONT SF9 workbook: {str(e)}")

    headers = {
        "Content-Disposition": f'attachment; filename="{out_filename}"',
        "X-Suggested-Filename": out_filename,
        "X-Student-Count": str(len(students)),
        "Access-Control-Expose-Headers": "Content-Disposition, X-Suggested-Filename, X-Student-Count"
    }

    return StreamingResponse(
        output_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers
    )

@app.post("/api/convert/back")
async def convert_back(
    file: UploadFile = File(...),
    overrides: Optional[str] = Form(None)
):
    """Generates and streams the BACK SF9 workbook (1 sheet per student)."""
    if not file.filename.endswith(('.xlsx', '.xlsm', '.xltx')):
        raise HTTPException(status_code=400, detail="Please upload a valid Excel file (.xlsx).")

    contents = await file.read()
    students = parse_input_workbook(contents)
    if not students:
        raise HTTPException(status_code=400, detail="No student records found.")

    global_overrides = {}
    if overrides:
        try:
            global_overrides = json.loads(overrides)
        except Exception:
            pass

    try:
        wb_back, out_filename = generate_sf9_back_workbook(students, global_overrides=global_overrides)
        output_stream = io.BytesIO()
        wb_back.save(output_stream)
        output_stream.seek(0)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate BACK SF9 workbook: {str(e)}")

    headers = {
        "Content-Disposition": f'attachment; filename="{out_filename}"',
        "X-Suggested-Filename": out_filename,
        "X-Student-Count": str(len(students)),
        "Access-Control-Expose-Headers": "Content-Disposition, X-Suggested-Filename, X-Student-Count"
    }

    return StreamingResponse(
        output_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers
    )

@app.post("/api/convert/zip")
@app.post("/api/convert")
async def convert_zip_or_unified(
    file: UploadFile = File(...),
    overrides: Optional[str] = Form(None),
    type: Optional[str] = Query("zip")
):
    """
    Generates and streams:
    - type=front -> FRONT workbook
    - type=back  -> BACK workbook
    - type=zip   -> Complete ZIP bundle with both workbooks
    """
    if type == "front":
        return await convert_front(file, overrides)
    elif type == "back":
        return await convert_back(file, overrides)

    # Default to ZIP bundle containing both FRONT and BACK
    if not file.filename.endswith(('.xlsx', '.xlsm', '.xltx')):
        raise HTTPException(status_code=400, detail="Please upload a valid Excel file (.xlsx).")

    contents = await file.read()
    students = parse_input_workbook(contents)
    if not students:
        raise HTTPException(status_code=400, detail="No student records found.")

    global_overrides = {}
    if overrides:
        try:
            global_overrides = json.loads(overrides)
        except Exception:
            pass

    try:
        zip_stream, zip_filename = generate_sf9_zip_bundle(students, global_overrides=global_overrides)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate SF9 bundle: {str(e)}")

    headers = {
        "Content-Disposition": f'attachment; filename="{zip_filename}"',
        "X-Suggested-Filename": zip_filename,
        "X-Student-Count": str(len(students)),
        "Access-Control-Expose-Headers": "Content-Disposition, X-Suggested-Filename, X-Student-Count"
    }

    return StreamingResponse(
        zip_stream,
        media_type="application/zip",
        headers=headers
    )

@app.get("/api/template")
async def download_template():
    if os.path.exists(INPUT_TEMPLATE_PATH):
        return FileResponse(
            INPUT_TEMPLATE_PATH,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="SF9_Data_Input_Template.xlsx"
        )
    raise HTTPException(status_code=404, detail="Template file not found.")

@app.get("/api/health")
async def health_check():
    return {
        "status": "ok",
        "front_template_exists": os.path.exists(FRONT_TEMPLATE_PATH),
        "back_template_exists": os.path.exists(BACK_TEMPLATE_PATH)
    }

if os.path.exists(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/")
async def get_index():
    index_file = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    return HTMLResponse("<h1>SF9 Converter Backend Running</h1><p>Static frontend not found.</p>")

if __name__ == "__main__":
    port = 8000
    print(f"==================================================")
    print(f"  SF9 AUTOMATED REPORT CARD GENERATOR (2026-2027)")
    print(f"  Running on: http://localhost:{port}")
    print(f"==================================================")
    try:
        webbrowser.open(f"http://localhost:{port}")
    except Exception:
        pass
    uvicorn.run(app, host="127.0.0.1", port=port)
