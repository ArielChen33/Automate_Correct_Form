from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import os
from backend.processor import process_excel
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

# Add middleware to avoid CORS issues
app.add_middleware(
    CORSMiddleware, 
    allow_origins = ["http://localhost:3000"], # or ["*"] for dev
    allow_credentials = True, 
    allow_methods = ["*"], 
    allow_headers = ["*"]
)

@app.get("/")
def homePage():
    return "Welcome! This is from the backend"

@app.get("/test")
def testing():
    # name = "Ariel"
    # message = f"Hi {name}, testing successful!" 
    return {"message": "welcome~"}

@app.post("/process/")
async def process_file(file: UploadFile = File(...)):
    # Ensure upload directory exists
    # os.makedirs("uploads", exist_ok=True)
    # os.makedirs("outputs", exist_ok=True)

    input_path = f"uploads/{file.filename}"
    with open(input_path, "wb") as f: 
        f.write(await file.read())
    
    output_path = process_excel(input_path)

    return FileResponse(output_path, filename=os.path.basename(output_path))

