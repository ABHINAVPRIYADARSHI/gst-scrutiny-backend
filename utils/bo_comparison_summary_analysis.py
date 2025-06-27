import os
from glob import glob
import pandas as pd
from utils.globals.constants import result_point_1, result_point_4, result_point_6, result_point_2

start_row_skip_nine_rows = 9  # Row 10 in Excel is index 9 in pandas
start_row_skip_six_rows = 6  # Row 7 in Excel is index 6 in pandas
sheet_tax_liability_summary = "Tax Liability Summary"
sheet_comparison_summary = "Comparison Summary"
sheet_reverse_charge = "Reverse charge"
sheet_ITC_Other_than_IMPG = "ITC (Other than IMPG)"
sheet_ITC_IMPG = "ITC (IMPG)"


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

        # Read Tax Liability sheet and draw analysis
        df_raw_tax_liability_sheet = pd.read_excel(bo_file, sheet_name=sheet_tax_liability_summary, header=None, skiprows=start_row_skip_nine_rows)
        df_raw_tax_liability_sheet = df_raw_tax_liability_sheet[df_raw_tax_liability_sheet.iloc[:, 0].notna()]  # Drop empty rows
        total_row_tax_liability_sheet = df_raw_tax_liability_sheet[df_raw_tax_liability_sheet.iloc[:, 0].astype(str).str.strip() == "Total"]
        if not total_row_tax_liability_sheet.empty:
            final_result_points["result_point_1_IGST"] = total_row_tax_liability_sheet.iloc[0, 21]  # Column V (index 21)
            final_result_points["result_point_1_CGST"] = total_row_tax_liability_sheet.iloc[0, 22]  # Column W (index 22)
            final_result_points["result_point_1_SGST"] = total_row_tax_liability_sheet.iloc[0, 23]  # Column X (index 23)
            final_result_points["result_point_1_CESS"] = total_row_tax_liability_sheet.iloc[0, 24]  # Column Y (index 24)
            # print(f"✅ Comparison Summary sheet: Total sum of GSTR-1 vs 3B shortfall: {shortfall_total_col_D}")

        # Read ITC(Other than IMPG) sheet and draw analysis
        df_raw_ITC_Other_IMPG_sheet = pd.read_excel(bo_file, sheet_name=sheet_ITC_Other_than_IMPG, header=None,
                                                    skiprows=start_row_skip_six_rows)
        df_raw_ITC_Other_IMPG_sheet = df_raw_ITC_Other_IMPG_sheet[
            df_raw_ITC_Other_IMPG_sheet.iloc[:, 0].notna()]  # Drop empty rows
        total_row_ITC_Other_IMPG_sheet = df_raw_ITC_Other_IMPG_sheet[
            df_raw_ITC_Other_IMPG_sheet.iloc[:, 0].astype(str).str.strip() == "Total"]

        # Read ITC(IMPG) sheet and draw analysis
        df_raw_ITC_IMPG_sheet = pd.read_excel(bo_file, sheet_name=sheet_ITC_IMPG, header=None,
                                                    skiprows=start_row_skip_six_rows)
        df_raw_ITC_IMPG_sheet = df_raw_ITC_IMPG_sheet[
            df_raw_ITC_IMPG_sheet.iloc[:, 0].notna()]  # Drop empty rows
        total_row_ITC_IMPG_sheet = df_raw_ITC_IMPG_sheet[
            df_raw_ITC_IMPG_sheet.iloc[:, 0].astype(str).str.strip() == "Total"]

        # IGST is Jth column
        igst = (pd.to_numeric(total_row_ITC_Other_IMPG_sheet.iloc[0, 9], errors='coerce') or 0)
        + (pd.to_numeric(total_row_ITC_IMPG_sheet.iloc[0, 5], errors='coerce') or 0)
        cgst = pd.to_numeric(total_row_ITC_Other_IMPG_sheet.iloc[0, 10], errors='coerce') or 0
        sgst = pd.to_numeric(total_row_ITC_Other_IMPG_sheet.iloc[0, 11], errors='coerce') or 0
        cess = (pd.to_numeric(total_row_ITC_Other_IMPG_sheet.iloc[0, 12], errors='coerce')or 0)
        + (pd.to_numeric(total_row_ITC_IMPG_sheet.iloc[0, 6], errors='coerce') or 0)
        final_result_points['result_point_4_IGST'] = igst
        final_result_points['result_point_4_CGST'] = cgst
        final_result_points['result_point_4_SGST'] = sgst
        final_result_points['result_point_4_CESS'] = cess

        # Read Comparison Summary sheet and draw analysis
        df_raw_CS_sheet = pd.read_excel(bo_file, sheet_name=sheet_comparison_summary, header=None,
                                        skiprows=start_row_skip_nine_rows)
        df_raw_CS_sheet = df_raw_CS_sheet[df_raw_CS_sheet.iloc[:, 0].notna()]  # Drop empty rows
        total_row = df_raw_CS_sheet[df_raw_CS_sheet.iloc[:, 0].astype(str).str.strip() == "Total"]
        if not total_row.empty:
            col_B_total = total_row.iloc[0, 1]  # Column B (index 1)
            final_result_points[result_point_6] = col_B_total
            print(f"Comparison Summary sheet: Total sum of column B: {col_B_total} ")
        else:
            print("❌ Comparison Summary sheet: Total row not found")

        # Read Reverse charge sheet and draw analysis
        df_raw_RC_sheet = pd.read_excel(bo_file, sheet_name=sheet_reverse_charge, header=None, skiprows=start_row_skip_six_rows)
        total_row_RC = df_raw_RC_sheet[df_raw_RC_sheet.iloc[:, 0].astype(str).str.strip() == "Total"]
        if not total_row_RC.empty:
            # Column letters J, K, L, M correspond to indices 9, 10, 11, 12 (0-based)
            final_result_points["result_point_2_IGST"] = total_row_RC.iloc[0, 9]
            final_result_points["result_point_2_CGST"] = total_row_RC.iloc[0, 10]
            final_result_points["result_point_2_SGST"] = total_row_RC.iloc[0, 11]
            final_result_points["result_point_2_CESS"] = total_row_RC.iloc[0, 12]
        else:
            print("❌ Reverse charge sheet: Total row not found")

        # === SAVE RESULTS to Excel ===


        # Split points into two sheets
        # df_comparison = pd.DataFrame([
        #     {'Description': "shortfall_total_col_D", 'Value': shortfall_total_col_D},
        #     {'Description': "shortfall_total_col_I", 'Value': shortfall_total_col_I},
        #     {'Description': "col_B_total", 'Value': col_B_total}
        # ])
        #
        # df_reverse_charge = pd.DataFrame([
        #     {'Description': 'total_sum_col_J_K_L_M', 'Value': total_sum_col_J_K_L_M}
        # ])
        #
        # # Output file
        # os.makedirs(output_path, exist_ok=True)
        # output_file = os.path.join(output_path, "BO_comparison_summary_analysis.xlsx")
        #
        # # Write to Excel with two sheets
        # with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        #     df_comparison.to_excel(writer, sheet_name="Comparison Summary", index=False)
        #     df_reverse_charge.to_excel(writer, sheet_name="Reverse charge", index=False)
        #
        #     # Get workbook and worksheet objects
        #     workbook = writer.book
        #     worksheet_comparison = writer.sheets["Comparison Summary"]
        #     worksheet_reverse = writer.sheets["Reverse charge"]
        #
        #     # Define a format with word wrap
        #     wrap_format = workbook.add_format({'text_wrap': True})
        #
        #     # Set column widths and apply format
        #     worksheet_comparison.set_column('A:A', 25, wrap_format)  # Column A = Description
        #     worksheet_comparison.set_column('B:B', 25, wrap_format)  # Column B = Value
        #
        #     worksheet_reverse.set_column('A:A', 40, wrap_format)
        #     worksheet_reverse.set_column('B:B', 20, wrap_format)
        # print(f"[BO comparison Analysis] ✅ Result saved to: {output_file}")
        return final_result_points
    except Exception as e:
        print(f"[BO comparison Analysis] ❌ Error during analysis: {e}")
        return final_result_points
