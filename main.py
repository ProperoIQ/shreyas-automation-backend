import shutil
import os
from fastapi import FastAPI, Query, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from functions.remove_decimals import remove_decimals_from_excel
import uvicorn
from fastapi import BackgroundTasks

from functions.segregator import process_multiple_files
from functions.consolidater import process_and_merge_files
from functions.balance_summary import process_file
from functions.age_summary import generate_summary
from functions.combiner import combine_sheets
from functions.get_details import fetch_all_reports
from functions.adjust_column_cells import process_output_folder , move_xlsx_to_output

app = FastAPI()

# ✅ Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://automation-frontend-589889616484.asia-south1.run.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)



def cleanup_output_folder(zip_filename: str):
    try:
        # Remove the zip file
        if os.path.exists(zip_filename):
            os.remove(zip_filename)

        # Remove all files in the output folder
        output_dir = 'output'
        for file in os.listdir(output_dir):
            file_path = os.path.join(output_dir, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
    except Exception as e:
        print(f"[Cleanup Error] {e}")

@app.post("/process_and_download")
async def process_and_download(
    background_tasks: BackgroundTasks,
    date_filter: str = Query(..., description="Date filter for fetching reports")
):
    try:
        # Your existing processing logic
        fetch_all_reports(date_filter)
        process_multiple_files('csvdata/input_invoice_aging_details_nvb.xlsx', 'csvdata/input_invoice_aging_details_smcs.xlsx')
        generate_summary(
            {
                'SMCS': 'output/SMCS_Age_Range_Columns.xlsx',
                'NVB': 'output/NVB_Age_Range_Columns.xlsx'
            },
            'output/Age_summary.xlsx'
        )
        process_and_merge_files(
            'csvdata/input_customer_balance_summary_details_nvb.xlsx',
            'csvdata/input_customer_balance_summary_details_smcs.xlsx',
            'output/unified_file.csv'
        )
        process_file('output/unified_file.xlsx', 'output/balances_summary.xlsx')
        combine_sheets('output/balances_summary.xlsx', 'output/Age_summary.xlsx', 'output/Final.xlsx')
        move_xlsx_to_output("csvdata", "output")
        process_output_folder()
        
        files_to_process = [
            'output/SMCS_Age_Range_Columns.xlsx',
            'output/NVB_Age_Range_Columns.xlsx',
            'output/Age_summary.xlsx',
            'output/unified_file.xlsx',
            'output/balances_summary.xlsx',
            'output/Final.xlsx'
        ]

        # Remove decimals from each file
        for file_path in files_to_process:
            if os.path.exists(file_path):
                remove_decimals_from_excel(file_path)

        # Zip the output folder
        zip_filename = 'output.zip'
        shutil.make_archive(zip_filename.replace('.zip', ''), 'zip', 'output')

        # Return file as response
        return FileResponse(zip_filename, filename="output.zip", media_type='application/zip')

    except FileNotFoundError as fnf_error:
        raise HTTPException(status_code=404, detail=f"File not found: {str(fnf_error)}")
    except PermissionError as perm_error:
        raise HTTPException(status_code=403, detail=f"Permission error: {str(perm_error)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"An unexpected error occurred: {str(e)}")
