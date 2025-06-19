import os
import pandas as pd
from utils.globals.constants import OLD_TABLE_POSITIONS_GSTR_3B
from utils.globals.constants import NEW_TABLE_POSITIONS_GSTR_3B
from utils.globals.constants import newFormat
from utils.globals.constants import estimatedITCReversal
from utils.globals.constants import diffInRCMPayment
from utils.globals.constants import diffInRCM_ITC
from utils.globals.constants import extract_table_with_header


async def gstr3b_merged_reader(gstin):
    print(" === Starting execution of file gstr3b_merged_reader.py ===")
    input_path = f"reports/{gstin}/GSTR-3B_merged.xlsx"
    try:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"[GSTR-3B reader]: Input file not found at {input_path}")
        else:
            valuesFrom3b = {}
            # Load full sheet without header
            df_full = pd.read_excel(input_path, sheet_name="GSTR-3B_merged", header=None)
            print("GSTR-3B_merged.xlsx fetched successfully.")
            # Read cell a1 (row 0, column 0)
            gstr3b_format = str(df_full.iat[0, 0]).strip()
            if gstr3b_format == newFormat:
                TABLE_POSITIONS = NEW_TABLE_POSITIONS_GSTR_3B
            else:
                TABLE_POSITIONS = OLD_TABLE_POSITIONS_GSTR_3B
            print(f"GSTR-3B_merged.xlsx is based on {gstr3b_format}")
            # Extract all tables
            tables = {}
            for key in TABLE_POSITIONS:
                tables[key] = extract_table_with_header(df_full, key, TABLE_POSITIONS)

            # 1. Estimated Proportionate ITC reversal calculation
            table_3_1 = tables["3.1"].copy()
            # print(f"Printing table 3.1: {table_3_1}")
            table_4 = tables["4"].copy()
            # print(table_4)
            table_6_1 = tables["6.1"].copy()

            # Sum of Table 4 (Part A Rows 1+2+4+5) = Sum of Total ITC availed (CGST + SGST + IGST + Cess)
            sum_table_4A_row_1 = pd.to_numeric(table_4.iloc[1, 1:], errors='coerce').sum(skipna=True)
            valuesFrom3b["sum_table_4A_row_1"] = sum_table_4A_row_1
            sum_table_4A_row_2 = pd.to_numeric(table_4.iloc[2, 1:], errors='coerce').sum(skipna=True)
            sum_table_4A_row_4 = pd.to_numeric(table_4.iloc[4, 1:], errors='coerce').sum(skipna=True)
            valuesFrom3b["sum_table_4A_row_4"] = sum_table_4A_row_4
            sum_table_4A_row_5 = pd.to_numeric(table_4.iloc[5, 1:], errors='coerce').sum(skipna=True)
            table_4_total_ITC = sum_table_4A_row_1 + sum_table_4A_row_2 + sum_table_4A_row_4 + sum_table_4A_row_5
            print(f"table_4_total_ITC (CGST + SGST + IGST + Cess): {table_4_total_ITC}")
            table_3_1_a1 = pd.to_numeric(table_3_1.iloc[0, 1], errors='coerce')  # 1st row
            table_3_1_B1 = pd.to_numeric(table_3_1.iloc[1, 1], errors='coerce')  # 2nd row
            table_3_1_C1 = pd.to_numeric(table_3_1.iloc[2, 1], errors='coerce')  # 3rd row
            table_3_1_D1 = pd.to_numeric(table_3_1.iloc[3, 1], errors='coerce')  # 3rd row
            table_3_1_E1 = pd.to_numeric(table_3_1.iloc[4, 1], errors='coerce')  # 5th row
            numerator = table_3_1_C1 + table_3_1_E1
            denominator = table_3_1_a1 + table_3_1_B1 + table_3_1_C1 + table_3_1_E1
            estimatedProITCReversal = round(((numerator / denominator if denominator else 0) * table_4_total_ITC), 2)
            valuesFrom3b[estimatedITCReversal] = estimatedProITCReversal
            print(f"Estimated Proportionate ITC Reversal calculation done: {estimatedProITCReversal}")
            valuesFrom3b["table_3_1_a+c+e_taxable_value"] = table_3_1_a1 + table_3_1_C1 + table_3_1_E1

            # 2. Difference of [table 3.1 (d)] -  [ table 4 A (3)] (CGST + SGST + IGST + Cess)
            # table 3.1 sum of row 4, 3rd column onwards : Inward supplies liable to reverse charge -
            # Table 4 A(3) Inward supplies liable to reverse charge (other than 1 & 2 above)
            sum_table_3_1_row_d = pd.to_numeric(table_3_1.iloc[3, 2:], errors='coerce').sum(skipna=True)
            sum_table_4A_row_3 = pd.to_numeric(table_4.iloc[3, 1:], errors='coerce').sum(skipna=True)
            diifIn_RCM_ITC = sum_table_3_1_row_d - sum_table_4A_row_3
            valuesFrom3b[diffInRCM_ITC] = diifIn_RCM_ITC
            print(f"Difference in RCM ITC: {diifIn_RCM_ITC}")

            # 3. Calculation : Difference in RCM payment
            # Difference of 3.1(d)] as above -  [table 6.1 B(Total tax payable)]
            # Get last 4 rows from column 1 of Table 6.1 of GSTR-3B_merged.xlsx
            # sum_row_d_table_3_1 = pd.to_numeric(table_3_1.iloc[3, 2:], errors='coerce').sum(skipna=True)
            sum_table6_column1_last_4_rows = pd.to_numeric(table_6_1.iloc[-4:, 1], errors='coerce').sum(skipna=True)
            diffIn_RCM_Pay = sum_table_3_1_row_d - sum_table6_column1_last_4_rows
            print(f"Difference in RCM payment: {diffIn_RCM_Pay}")
            valuesFrom3b[diffInRCMPayment] = diffIn_RCM_Pay
            print(f"Difference in RCM payment done: {diffInRCMPayment}")

            # 4. Total tax on Inward supplies (liable to reverse charge) table_3_1_D1 (CGST + SGST + IGST + Cess)
            table_3_1_D1_text = "Total tax on" + table_3_1.iloc[3, 0]
            table_3_1_D2 = sum_table_3_1_row_d  # Sum of row D
            valuesFrom3b["table_3_1_D1"] = table_3_1_D1_text
            valuesFrom3b["table_3_1_D2"] = table_3_1_D2
            print(f"Total tax on Inward supplies (liable to reverse charge) (CGST + SGST + IGST + Cess): {table_3_1_D2}")

            # 5. Sum of Table 3.1 a + b + d (CGST + SGST + IGST + Cess)
            sum_table_3_1_row_a = pd.to_numeric(table_3_1.iloc[0, 2:], errors='coerce').sum(skipna=True)
            sum_table_3_1_row_b = pd.to_numeric(table_3_1.iloc[1, 2:], errors='coerce').sum(skipna=True)
            sum_table_3_1_row_d = pd.to_numeric(table_3_1.iloc[3, 2:], errors='coerce').sum(skipna=True)
            sum_table_3_1_row_a_b_d_taxes = sum_table_3_1_row_a + sum_table_3_1_row_b + sum_table_3_1_row_d
            valuesFrom3b["sum_table_3_1_row_a_b_d_taxes"] = sum_table_3_1_row_a_b_d_taxes
            print(f"Sum 3.1 a + b + d (CGST + SGST + IGST + Cess): {sum_table_3_1_row_a_b_d_taxes}")

            # 6. Table 3.1 (C1 +E1)
            table_3_1_C1_text = "Taxable value" + table_3_1.iloc[2, 0] + "And" + table_3_1.iloc[4, 0]
            table_3_1_C1_plus_E1 = table_3_1.iloc[2, 1] + table_3_1.iloc[4, 1]
            valuesFrom3b["table_3_1_C1"] = table_3_1_C1_text
            valuesFrom3b["table_3_1_C2"] = table_3_1_C1_plus_E1
            print(f"Table 3.1 (C1 +E1) Taxable value: {table_3_1_C1_plus_E1}")

            # 7. Table 3.1 (a1+B1+D1+E1-C1)
            sum_table_3_1_A1_B1_D1_E1_minus_C1 = table_3_1_a1 + table_3_1_B1 + table_3_1_D1 + table_3_1_E1 - table_3_1_C1
            valuesFrom3b["sum_table_3_1_A1_B1_D1_E1_minus_C1"] = sum_table_3_1_A1_B1_D1_E1_minus_C1
            print(f"Table 3.1 (A1+B1+D1+E1-C1) calculation done: {sum_table_3_1_A1_B1_D1_E1_minus_C1}")

            # 8. We need only sum of Total Tax Payable (column 1) of Table 6.1 of GSTR-3B_merged.xlsx
            total_tax_payable_column_GSTR_3B_table_6 = pd.to_numeric(table_6_1.iloc[:, 1], errors='coerce').sum(skipna=True)
            valuesFrom3b["total_tax_payable_column_GSTR_3B_table_6"] = total_tax_payable_column_GSTR_3B_table_6

            print(" === ✅ Returning after successful execution of file gstr3b_merged_reader.py ===")
            return valuesFrom3b
    except Exception as e:
        print(f"[GSTR-3B_reader] ❌ Error: {e}")
        raise Exception(e)
