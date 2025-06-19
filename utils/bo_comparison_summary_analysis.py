import os
from glob import glob
import pandas as pd
from utils.globals.constants import result_point_1, result_point_4, result_point_6, result_point_2

start_row_comparison_summary_sheet = 9  # Row 10 in Excel is index 9 in pandas
start_row_reverse_charge_sheet = 6  # Row 7 in Excel is index 6 in pandas


async def generate_bo_comparison_summary_analysis(gstin):
    print(f" === Started execution of method generate_bo_comparison_summary_analysis for: {gstin} ===")
    final_result_points = {}
    input_path = f"uploaded_files/{gstin}/BO comparison summary/"
    output_path = f"reports/{gstin}/"

    try:
        bo_file_list = glob(os.path.join(input_path, "*.xlsx"))
        print(f"bo_file_list: {bo_file_list}")
        if not bo_file_list:
            raise FileNotFoundError(f"[BO comparison]: Input file not found at {input_path}")
        print(f"Found {len(bo_file_list)} BO comparison file(s).")
        # Load Excel and extract headers from first file only
        bo_file = bo_file_list[0]

        # Read Comparison Summary sheet and draw analysis
        df_raw_CS_sheet = pd.read_excel(bo_file, sheet_name="Comparison Summary", header=None, skiprows=start_row_comparison_summary_sheet)
        df_raw_CS_sheet = df_raw_CS_sheet[df_raw_CS_sheet.iloc[:, 0].notna()]  # Drop empty rows
        total_row = df_raw_CS_sheet[df_raw_CS_sheet.iloc[:, 0].astype(str).str.strip() == "Total"]
        if not total_row.empty:
            shortfall_total_col_D = total_row.iloc[0, 3]  # Column D (index 3)
            final_result_points[result_point_1] = shortfall_total_col_D
            print(f"✅ Comparison Summary sheet: Total sum of GSTR-1 vs 3B shortfall:: {shortfall_total_col_D}")
            shortfall_total_col_I = total_row.iloc[0, 8]  # Column I (index 8)
            final_result_points[result_point_4] = shortfall_total_col_I
            print(f"Comparison Summary sheet: Total sum of GSTR-2B vs 3B shortfall: {shortfall_total_col_I} ")
            col_B_total = total_row.iloc[0, 1]  # Column B (index 1)
            final_result_points[result_point_6] = col_B_total
            print(f"Comparison Summary sheet: Total sum of column B: {col_B_total} ")
        else:
            print("❌ Comparison Summary sheet: Total row not found")

        # Read Reverse charge sheet and draw analysis
        df_raw_RC_sheet = pd.read_excel(bo_file, sheet_name="Reverse charge", header=None, skiprows=start_row_reverse_charge_sheet)
        total_row_RC = df_raw_RC_sheet[df_raw_RC_sheet.iloc[:, 0].astype(str).str.strip() == "Total"]
        if not total_row_RC.empty:
            # Column letters J, K, L, M correspond to indices 9, 10, 11, 12 (0-based)
            total_sum = pd.to_numeric(total_row_RC.iloc[0, 9:13], errors="coerce").sum()
            final_result_points[result_point_2] = total_sum
            print(f"Reverse charge sheet  columns J+K+L+M sum: {total_sum} ")
        else:
            print("❌ Reverse charge sheet: Total row not found")
        return final_result_points
    except Exception as e:
        print(f"[BO comparison Analysis] ❌ Error during analysis: {e}")
        return final_result_points
