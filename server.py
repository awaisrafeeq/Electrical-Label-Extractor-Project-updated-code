from pathlib import Path
import time
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from extract_equipment_simple import main  # Ensure that the extraction function is correct
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
# ... (existing code)

app = FastAPI()

# Allow the HTML page to call this API from any origin
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Restrict this later to your specific domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.mount("/outputs", StaticFiles(directory="outputs"), name="outputs")
app.mount("/", StaticFiles(directory="static", html=True), name="static")
# Folder where Excel files will be saved
OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

# Serve the Excel files so the frontend can download them
app.mount("/outputs", StaticFiles(directory=str(OUTPUT_DIR)), name="outputs")


@app.post("/extract")
async def extract_equipment(file: UploadFile = File(...)):
    """
    Accepts a PDF upload from the frontend, runs the extraction process,
    saves an Excel file, and returns counts plus a download URL.
    """
    if file.content_type != "application/pdf":
        return JSONResponse(
            status_code=400,
            content={"detail": "Only PDF files are allowed"},
        )

    # Unique file name for the output Excel file
    timestamp = int(time.time())
    output_name = f"equipment_data_{timestamp}.xlsx"
    output_path = OUTPUT_DIR / output_name

    try:
        # Call the existing main function from the extraction logic
        df, equipment_data = main(file.file, str(output_path))

        if df is None:
            return JSONResponse(
                status_code=500,
                content={"detail": "No equipment found in PDF"},
            )

        # Compute counts for MVS and DSG types from the DataFrame
        mvs_count = int((df["Type"] == "MVS").sum()) if "Type" in df.columns else 0
        dsg_count = int((df["Type"] == "DSG").sum()) if "Type" in df.columns else 0

        # Prepare the URL for downloading the generated Excel file
        excel_url = f"/outputs/{output_name}"

        # Prepare the response with the counts and file download link
        return {
            "mvs_count": mvs_count,
            "dsg_count": dsg_count,
            "excel_url": excel_url,
            "output_name": output_name,
            "equipment_list": equipment_data,  # Returning the extracted equipment data as well
        }

    except Exception as e:
        # In case of any error, return a failure response
        return JSONResponse(
            status_code=500,
            content={"detail": f"Extraction failed. {str(e)}"},
        )
