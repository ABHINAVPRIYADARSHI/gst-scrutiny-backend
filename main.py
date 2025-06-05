import os, glob, logging
from http.client import HTTPException
from typing import List
from fastapi import FastAPI, UploadFile, File, Form, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from utils.file_handler import save_uploaded_file
from utils.pdf_processor import process_pdf_files
from utils.csv_processor import process_csv_files
from utils.master_generator import generate_master_excel_for_return_type
return_types = ["GSTR-1", "GSTR-2A", "GSTR-3B", "GSTR-9"]  # Add more return types when needed

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    # allow_origins=["http://localhost:3000"],  # Frontend origin
    allow_origins=["*"],  # Frontend origin
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload/")
async def upload_files(
    gstn: str = Form(...),
    return_type: str = Form(...),
    files: List[UploadFile] = File(...)
):
    saved_paths = []
    for file in files:
        path = await save_uploaded_file(file, gstn=gstn, return_type=return_type)
        saved_paths.append(path)
    return {"file_paths": saved_paths}

@app.get("/files/")
def list_uploaded_files(gstn: str, return_type: str):
    folder_path = os.path.join("uploaded_files", gstn, return_type)
    if not os.path.exists(folder_path):
        return {"files": []}
    
    files = [os.path.basename(path) for path in glob.glob(f"{folder_path}/*")]
    return {"files": files}

@app.post("/process/")
async def process_files(return_type: str, file_paths: List[str]):
    if all(path.endswith(".pdf") for path in file_paths):
        output_path = process_pdf_files(file_paths, return_type)
    elif all(path.endswith(".csv") for path in file_paths):
        output_path = process_csv_files(file_paths, return_type)
    else:
        return {"error": "Mix of CSV and PDF not supported."}
    return {"output": output_path}

@app.get("/reports/{filename}")
def get_report_file(filename: str):
    file_path = os.path.join("reports", filename)
    return FileResponse(file_path, media_type="application/octet-stream", filename=filename)

@app.delete("/delete/")
def delete_file(gstn: str, return_type: str, filename: str):
    folder_path = os.path.join("uploaded_files", gstn, return_type)
    file_path = os.path.join(folder_path, filename)
    if os.path.exists(file_path):
        os.remove(file_path)
        return {"message": "File deleted successfully."}
    return JSONResponse(status_code=404, content={"error": "File not found."})

@app.post("/generate_master/")
async def generate_master(gstn: str = Form(...)):
    """Generate master Excel files for all return types for a given GSTIN."""
    generated_reports = []
    for rt in return_types:
        input_dir = f"uploaded_files/{gstn}/{rt}"
        output_dir = f"reports/{gstn}/{rt}"
        os.makedirs(output_dir, exist_ok=True)

        if not os.path.exists(input_dir) or not os.listdir(input_dir):
            print(f"[{rt}] Skipped: No input files in {input_dir}")
            continue

        try:
            output_file = await generate_master_excel_for_return_type(rt, input_dir, output_dir)
            generated_reports.append({"return_type": rt, "report": output_file})
        except Exception as e:
            print(f"[{rt}] Error: {e}")
            continue

    if not generated_reports:
        raise HTTPException(status_code=404, detail="No reports generated for any return type")

    return JSONResponse(content={"status": "completed", "reports": generated_reports})
