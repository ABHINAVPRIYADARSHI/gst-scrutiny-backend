import os
import pandas as pd
from utils.globals.constants import ewb_out_MIS_report, result_point_6

TAX_COL = [7, 8]  # Assess val. and Tax val.
sheet_name = ["By_HSN_Code", "By_GSTIN", "By_Vehicle_No."]


async def generate_ewb_out_merged_analysis(gstin):
    print(" === Starting execution of file ewb_out_merged_analysis.py ===")
    input_path = f"reports/{gstin}/EWB-Out_merged.xlsx"
    output_path = f"reports/{gstin}/EWB-Out_merged_analysis.xlsx"
    final_result_points = {}
    df_by_hsn_code = pd.DataFrame()
    df_by_GSTIN = pd.DataFrame()
    df_by_vehicle_no = pd.DataFrame()
    try:
        if not os.path.exists(input_path):
            print(f"[EWB-Out_merged analysis] Skipped: Input file not found at {input_path}")
            return final_result_points

        df = pd.read_excel(input_path, sheet_name=ewb_out_MIS_report, header=0)
        # Convert relevant tax columns to numeric
        for col in TAX_COL:
            df.iloc[:, col] = pd.to_numeric(df.iloc[:, col], errors='coerce')

        #  1. Group by HSN
        try:
            df_by_hsn_code = df.groupby('HSN Code')[['Assess Val.', 'Tax Val.']].sum().reset_index()
            total_row = df_by_hsn_code[['Assess Val.', 'Tax Val.']].sum()
            total_row['HSN Code'] = 'Total'
            df_by_hsn_code = pd.concat([df_by_hsn_code, pd.DataFrame([total_row])], ignore_index=True)
        except Exception as e:
            print(f"[EWB-Out_merged_Analysis] ❌ Error while Group by HSN: {e}")

        # 2. Group by GSTIN
        # Step 1: Extract GSTIN from combined column
        try:
            df['GSTIN'] = df['To GSTIN & Name'].str.split('/').str[0].str.strip()
            df_by_GSTIN = df.groupby('GSTIN')[['Assess Val.', 'Tax Val.']].sum().reset_index()  # Step 2: Group by GSTIN
            total_row = df_by_GSTIN[['Assess Val.', 'Tax Val.']].sum()
            total_row['GSTIN'] = 'Total'
            df_by_GSTIN = pd.concat([df_by_GSTIN, pd.DataFrame([total_row])], ignore_index=True)
            total_tax_val = df_by_GSTIN["Tax Val."].iloc[-1]
            final_result_points[result_point_6] = total_tax_val
        except Exception as e:
            print(f"[EWB-Out_merged_Analysis] ❌ Error while group by Group by GSTIN: {e}")

        # 3. Group by Vehicle no.
        try:
            df_by_vehicle_no = df.groupby('Latest Vehicle No.')[['Assess Val.', 'Tax Val.']].sum().reset_index()
            total_row = df_by_vehicle_no[['Assess Val.', 'Tax Val.']].sum()
            total_row['Latest Vehicle No.'] = 'Total'
            df_by_vehicle_no = pd.concat([df_by_vehicle_no, pd.DataFrame([total_row])], ignore_index=True)
        except Exception as e:
            print(f"[EWB-Out_merged_Analysis] ❌ Error while group by Vehicle no: {e}")

        # Write each table to its own sheet
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_by_hsn_code.to_excel(writer, index=False, sheet_name="By_HSN_Code")
            df_by_GSTIN.to_excel(writer, index=False, sheet_name="By_GSTIN")
            df_by_vehicle_no.to_excel(writer, index=False, sheet_name="By_Vehicle_No.")

            workbook = writer.book  # ✅ Correct place to call add_format
            wrap_format = workbook.add_format({'text_wrap': True})  # Create a wrapped cell format
            for sheet in sheet_name:
                worksheet = writer.sheets[sheet]
                worksheet.set_column(0, 3, 25, wrap_format)

        print(f"[EWB-Out_merged_Analysis] ✅ Summary report generated at: {output_path}")
        return final_result_points
    except Exception as e:
        print(f"[EWB-Out_merged_analysis.py] ❌ Error: {e}")
        return final_result_points
        # raise Exception(e)
