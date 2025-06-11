import os
import pandas as pd
from glob import glob
from collections import defaultdict
from utils.extractors.gstr3b_table_extractor import extract_fixed_tables_from_gstr3b

manual_columns = [
                    "Description",
                    "Total Tax Payable",
                    "Integrated Tax paid through ITC",
                    "Central Tax paid through ITC",
                    "State/UT Tax paid through ITC",
                    "Cess paid through ITC",
                    "Tax paid in cash",
                    "Interest paid in cash",
                    "Late fee paid in cash"
                ]
async def generate_gstr3b_merged(input_dir, output_dir):
    print("Generating GSTR-3B merged report...")

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
        if key == "6.1":
            preprocess_table_6(df_list)
        base_df = df_list[0].copy(deep=True)
        # print(f"\nProcessing table: {key}")
        # print(f"Number of files: {len(df_list)}")
        for row_idx in range(base_df.shape[0]):
            for col_idx in range(1, base_df.shape[1]):
                total = 0.0
                for df_num, df in enumerate(df_list):
                    try:
                        val = df.iat[row_idx, col_idx]
                        # print(f"File {df_num}: Cell[{row_idx},{col_idx}] = {val}")
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
    output_path = os.path.join(output_dir, "GSTR-3B_merged.xlsx")

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        start_row = 0
        sheet_name = "GSTR-3B_merged"
        for key, df in final_tables.items():
            title_df = pd.DataFrame([[f"Table {key}"]])
            title_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False, header=False)
            start_row += 1
            df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
            start_row += len(df) + 2
        worksheet = writer.sheets[sheet_name]
        fixed_width = 40
        worksheet.set_column(0, 8, fixed_width)

    print(f"GSTR-3B_merged saved to: {output_path}")
    return output_path  # ✅ Return the file path for use in API response

def preprocess_table_6(df_list):
    """ Cleans and standardizes the structure of GSTR-3B Table 6.1 across multiple files.
        Assumes df_list is a list of DataFrames for Table 6.1.
        """
    for i in range(len(df_list)):
        df_list[i].columns = manual_columns
        df_list[i] = df_list[i].iloc[1:].reset_index(drop=True)
    return df_list

