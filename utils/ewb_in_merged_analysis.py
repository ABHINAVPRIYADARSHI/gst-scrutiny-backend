import os
import pandas as pd
from utils.globals.constants import ewb_in_MIS_report


TAX_COL = [7, 8]  # Assess val. and Tax val.
sheet_name = ["By_HSN_Code", "By_GSTIN", "By_Vehicle_No."]


async def generate_ewb_in_merged_analysis(gstin):
    print(" === Starting execution of file ewb_in_merged_analysis.py ===")
    input_path = f"reports/{gstin}/EWB-In_merged.xlsx"
    output_path = f"reports/{gstin}/EWB-In_merged_analysis.xlsx"
    try:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"[EWB-In_merged analysis]: Input file not found at {input_path}")

        df = pd.read_excel(input_path, sheet_name=ewb_in_MIS_report, header=0)
        # Convert relevant tax columns to numeric
        for col in TAX_COL:
            df.iloc[:, col] = pd.to_numeric(df.iloc[:, col], errors='coerce')

        #  1. Group by HSN
        df_by_hsn_code = df.groupby('HSN Code')[['Assess Val.', 'Tax Val.']].sum().reset_index()
        total_row_hsn = df_by_hsn_code[['Assess Val.', 'Tax Val.']].sum()
        total_row_hsn['HSN Code'] = 'Total'
        df_by_hsn_code = pd.concat([df_by_hsn_code, pd.DataFrame([total_row_hsn])], ignore_index=True)

        # 2. Group by GSTIN
        # Step 1: Extract GSTIN from combined column
        df['GSTIN'] = df['From GSTIN & Name'].str.split('/').str[0].str.strip()
        df_by_GSTIN = df.groupby('GSTIN')[['Assess Val.', 'Tax Val.']].sum().reset_index()  # Step 2: Group by GSTIN
        total_row_GSTIN = df_by_GSTIN[['Assess Val.', 'Tax Val.']].sum()
        total_row_GSTIN['GSTIN'] = 'Total'
        df_by_GSTIN = pd.concat([df_by_GSTIN, pd.DataFrame([total_row_GSTIN])], ignore_index=True)

        # 3. Group by Vehicle no.
        df_by_vehicle_no = df.groupby('Latest Vehicle No.')[['Assess Val.', 'Tax Val.']].sum().reset_index()
        total_row_vehicle = df_by_vehicle_no[['Assess Val.', 'Tax Val.']].sum()
        total_row_vehicle['Latest Vehicle No.'] = 'Total'
        df_by_vehicle_no = pd.concat([df_by_vehicle_no, pd.DataFrame([total_row_vehicle])], ignore_index=True)

        # Write each table to its own sheet
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_by_hsn_code.to_excel(writer, index=False, sheet_name="By_HSN_Code")
            df_by_GSTIN.to_excel(writer, index=False, sheet_name="By_GSTIN")
            df_by_vehicle_no.to_excel(writer, index=False, sheet_name="By_Vehicle_No.")

            workbook = writer.book  # ✅ Correct place to call add_format
            # Create a wrapped cell format
            wrap_format = workbook.add_format({'text_wrap': True})
            for sheet in sheet_name:
                worksheet = writer.sheets[sheet]
                worksheet.set_column(0, 2, 25, wrap_format)

        print(f"[EWB-In_merged_Analysis] ✅ Summary report generated at: {output_path}")
        return output_path
    except Exception as e:
        print(f"[EWB-In_merged_analysis.py] ❌ Error: {e}")
        # raise Exception(e)

