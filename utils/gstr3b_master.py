import os
import pandas as pd
from glob import glob
from collections import defaultdict
from utils.extractors.gstr3b_table_extractor import extract_fixed_tables_from_gstr3b

async def generate_gstr3b_master(input_dir, output_dir):
    print("Generating GSTR-3B master report...")

    pdf_files = glob(os.path.join(input_dir, "*.pdf"))
    if not pdf_files:
        raise FileNotFoundError("No PDF files found in input directory.")

    print(f"Found {len(pdf_files)} PDF files.")

    combined_tables = defaultdict(list)

    for pdf_path in pdf_files:
        table_map = extract_fixed_tables_from_gstr3b(pdf_path)
        for key, df in table_map.items():
            combined_tables[key].append(df)

    final_tables = {}
    for key, df_list in combined_tables.items():
        base_df = df_list[0].copy(deep=True)

        print(f"\nProcessing table: {key}")
        print(f"Number of files: {len(df_list)}")
        for row_idx in range(base_df.shape[0]):
            for col_idx in range(1, base_df.shape[1]):
                total = 0.0
                for df_num, df in enumerate(df_list):
                    try:
                        val = df.iat[row_idx, col_idx]
                        print(f"File {df_num}: Cell[{row_idx},{col_idx}] = {val}")
                        num = pd.to_numeric(val, errors='coerce')
                        if pd.notnull(num):
                            total += num
                            # print(f"  File {df_num}: Cell[{row_idx},{col_idx}] = {val} → {num}")
                    except Exception as e:
                        print(f"  [Error] File {df_num}: Cell[{row_idx},{col_idx}] → {e}")
                        continue

                base_df.iat[row_idx, col_idx] = total #if pd.notnull(total) and total != 0 else ""

        final_tables[key] = base_df



    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "GSTR3B_Master_Report.xlsx")

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        start_row = 0
        sheet_name = "Summary"
        for key, df in final_tables.items():
            title_df = pd.DataFrame([[f"Table {key}"]])
            title_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False, header=False)
            start_row += 1
            df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
            start_row += len(df) + 2

    print(f"GSTR-3B master report saved to: {output_path}")
    return output_path  # ✅ Return the file path for use in API response

