import os
import pandas as pd
from glob import glob

from utils.globals.constants import ewb_in_MIS_report


async def generate_ewb_in_merged(input_dir, output_dir):
    print(f"[EWB-IN_merged.py] Started execution of method generate_ewb_in_merged for: {input_dir}")
    xls_files = glob(os.path.join(input_dir, '*.xls'))
    if not xls_files:
        print(f"[EWB-IN_merged.py] Skipped: Input file not found at: {input_dir}")
        return None
    print(f"[EWB-IN_merged.py] Found {len(xls_files)} Excel files to merge.")

    merged_df = pd.DataFrame()
    dataframes = []
    for file_path in sorted(xls_files):
        try:
            print(f"[EWB-IN_merged.py] Processing file: {file_path}")
            df = pd.read_html(file_path)[0]  # These are .html files disguised as .xls. We take 1st table.
            dataframes.append(df)
        except Exception as e:
            print(f"[EWB-IN_merged.py] ❌ Failed to read file {file_path}: {str(e)}")
    if dataframes:
        merged_df = pd.concat(dataframes, ignore_index=True)
        dataframes.clear()
    # Write merged DataFrame to .xlsx
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "EWB-In_merged.xlsx")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        merged_df.to_excel(writer, index=False, sheet_name=ewb_in_MIS_report)

        workbook = writer.book  # ✅ Correct place to call add_format
        wrap_format = workbook.add_format({'text_wrap': True})
        worksheet = writer.sheets[ewb_in_MIS_report]
        worksheet.set_column(0, 11, 20, wrap_format)  # From columns 0 to 12

    print(f"[EWB-IN_merged.py] Merged file saved: {output_path}")
    print("[EWB-IN_merged.py] Completed execution of method merge_ewb_in_files.")
    return output_dir
