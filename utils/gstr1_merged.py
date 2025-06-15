import os
import pandas as pd
from glob import glob
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook


async def generate_gstr1_merged(input_dir, output_dir):
    print(f"Started execution of method generate_gstr1_merged for: {input_dir}")
    excel_files = sorted(glob(os.path.join(input_dir, "*.xlsx")))
    if not excel_files:
        raise FileNotFoundError("No GSTR-1 Excel files found in input directory.")

    print(f"[GSTR-1] Found {len(excel_files)} Excel files.")

    header_map = {}  # sheet_name -> [header_row_1, header_row_2]
    sheet_data = defaultdict(list)

    for file_index, file_path in enumerate(excel_files):
        wb = load_workbook(file_path, data_only=True)
        print(f"Processing file: {os.path.basename(file_path)}")

        for sheet_name in wb.sheetnames:
            if sheet_name.lower().strip() == "read me":
                continue

            ws = wb[sheet_name]

            # Capture headers only from the first file
            if file_index == 0:
                header_rows = []
                for row_idx in [2, 3]:  # 3rd and 4th rows (0-indexed)
                    row_values = [cell.value if cell.value is not None else "" for cell in ws[row_idx + 1]]
                    header_rows.append(row_values)
                header_map[sheet_name] = header_rows

            # Extract data from row 5 onward
            data_rows = []
            for row in ws.iter_rows(min_row=5, values_only=True):
                if all(cell is None for cell in row):
                    continue
                data_rows.append(row)

            if data_rows:
                df = pd.DataFrame(data_rows)
                sheet_data[sheet_name].append(df)
            else:
                sheet_data[sheet_name]  # Ensure key exists even if empty

    # ✅ Write final output Excel
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "GSTR-1_merged.xlsx")
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, header_rows in header_map.items():
        ws = wb.create_sheet(title=(sheet_name + "_merged")[:31])  # Excel max length: 31 chars
        # Write both header rows
        for header in header_rows:
            ws.append(header)
        # Write data (if any)
        df_list = sheet_data.get(sheet_name)
        if df_list:
            combined_df = pd.concat(df_list, ignore_index=True)
            for row in dataframe_to_rows(combined_df, index=False, header=False):
                ws.append(row)

    wb.save(output_path)
    print(f"✅ GSTR-1 merged Excel saved to: {output_path}")
    return output_path
