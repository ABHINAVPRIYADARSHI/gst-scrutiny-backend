from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
import os
from typing import List

app = FastAPI()

# CORS for frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://localhost:5173"],  # Add Vite or React port
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "uploaded_files"
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.post("/upload/")
async def save_uploaded_file(file: UploadFile, gstn: str, return_type: str) -> str:
    folder = os.path.join("uploaded_files", gstn, return_type)
    os.makedirs(folder, exist_ok=True)
    file_path = os.path.join(folder, file.filename)
    
    with open(file_path, "wb") as f:
        f.write(await file.read())

    return file_path
