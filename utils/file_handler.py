from fastapi import FastAPI, UploadFile, File, Form
import os

app = FastAPI()

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
