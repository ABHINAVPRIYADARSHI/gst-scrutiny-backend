import os
import pandas as pd
from glob import glob
from utils.globals.constants import ewb_out_MIS_report


async def generate_ewb_out_merged(input_dir, output_dir):
    print(f" Started execution of method generate_ewb_out_merged for: {input_dir}")
    xlsx_files = glob(os.path.join(input_dir, '*.xlsx'))
    xls_files = glob(os.path.join(input_dir, '*.xls'))
    all_files = xlsx_files + xls_files
    if not all_files:
        raise FileNotFoundError("No EWB-OUT Excel files found in the input directory.")
    print(f"[EWB-Out_merged.py] Found {len(all_files)} Excel files to merge.")

    merged_df = None
    for idx, file_path in enumerate(sorted(all_files)):
        try:
            print(f"[INFO] Processing file: {file_path}")
            df = pd.read_html(file_path)[0]  # These are .html files disguised as .xls. We take 1st table.
            df.to_excel("converted.xlsx", index=False)
            # If first file → keep header + data
            if idx == 0:
                merged_df = df.copy()
            else:
                # Append only data rows (skip header)
                merged_df = pd.concat([merged_df, df], ignore_index=True)
        except Exception as e:
            print(f"[ERROR] Failed to read file {file_path}: {str(e)}")
            continue
    # Write merged DataFrame to .xlsx
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "EWB-Out_merged.xlsx")
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        merged_df.to_excel(writer, index=False, sheet_name=ewb_out_MIS_report)

        workbook = writer.book  # ✅ Correct place to call add_format
        wrap_format = workbook.add_format({'text_wrap': True})
        worksheet = writer.sheets[ewb_out_MIS_report]
        worksheet.set_column(0, 11, 20, wrap_format)  # From columns 0 to 12

    print(f"[EWB_Out_handler.py] Merged file saved: {output_path}")
    print("=== [EWB_Out_handler.py] Completed execution of method merge_ewb_out_files ===")
    return output_dir
