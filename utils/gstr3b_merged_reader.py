import os

import pandas as pd

from utils.globals.constants import NEW_TABLE_POSITIONS_GSTR_3B
from utils.globals.constants import OLD_TABLE_POSITIONS_GSTR_3B
from utils.globals.constants import extract_table_with_header
from utils.globals.constants import newFormat


async def gstr3b_merged_reader(gstin):
    print(f"[GSTR-3B_reader] Starting execution of file gstr3b_merged_reader.py ===")
    input_path = f"reports/{gstin}/GSTR-3B_merged.xlsx"
    valuesFrom3b = {}
    tables = {}
    try:
        if not os.path.exists(input_path):
            print(f"[GSTR-3B reader] Skipped: Input file not found at {input_path}")
            return valuesFrom3b
        else:
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
            for key in TABLE_POSITIONS:
                tables[key] = extract_table_with_header(df_full, key, TABLE_POSITIONS)

            # Taxpayer's info
            try:
                table_1 = tables["1"].copy()
                valuesFrom3b["financial_year"] = get_stripped_value(table_1.iloc[0, 1])
                table_2 = tables["2"].copy()
                valuesFrom3b["gstin_of_taxpayer"] = get_stripped_value(table_2.iloc[0, 1])
                valuesFrom3b["legal_name_of_taxpayer"] = get_stripped_value(table_2.iloc[1, 1])
                valuesFrom3b["trade_name_of_taxpayer"] = get_stripped_value(table_2.iloc[2, 1])
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting Taxpayer's info: {e}")

            try:
                # 1. Estimated Proportionate ITC reversal calculation
                table_3_1 = tables["3.1"].copy()
                # print(f"Printing table 3.1: {table_3_1}")
                table_4 = tables["4"].copy()
                # print(table_4)
                table_6_1 = tables["6.1"].copy()

                # Sum of Table 4 (Part A Rows 1+2+4+5) = Sum of Total ITC availed (CGST + SGST + IGST + Cess)
                sum_table_4A_IGST = pd.to_numeric(table_4.iloc[[1, 2, 4, 5], 1], errors='coerce').sum()
                sum_table_4A_CGST = pd.to_numeric(table_4.iloc[[1, 2, 4, 5], 2], errors='coerce').sum()
                sum_table_4A_SGST = pd.to_numeric(table_4.iloc[[1, 2, 4, 5], 3], errors='coerce').sum()
                sum_table_4A_CESS = pd.to_numeric(table_4.iloc[[1, 2, 4, 5], 4], errors='coerce').sum()

                table_3_1_a1 = pd.to_numeric(table_3_1.iloc[0, 1], errors='coerce')  # 1st row
                table_3_1_B1 = pd.to_numeric(table_3_1.iloc[1, 1], errors='coerce')  # 2nd row
                table_3_1_C1 = pd.to_numeric(table_3_1.iloc[2, 1], errors='coerce')  # 3rd row
                table_3_1_D1 = pd.to_numeric(table_3_1.iloc[3, 1], errors='coerce')  # 3rd row
                table_3_1_E1 = pd.to_numeric(table_3_1.iloc[4, 1], errors='coerce')  # 5th row
                numerator = table_3_1_C1 + table_3_1_E1
                denominator = table_3_1_a1 + table_3_1_B1 + table_3_1_C1 + table_3_1_E1
                valuesFrom3b["estimated_ITC_Reversal_IGST"] = round(((numerator / denominator if denominator else 0) * sum_table_4A_IGST), 2)
                valuesFrom3b["estimated_ITC_Reversal_CGST"] = round(((numerator / denominator if denominator else 0) * sum_table_4A_CGST), 2)
                valuesFrom3b["estimated_ITC_Reversal_SGST"] = round(((numerator / denominator if denominator else 0) * sum_table_4A_SGST), 2)
                valuesFrom3b["estimated_ITC_Reversal_CESS"] = round(((numerator / denominator if denominator else 0) * sum_table_4A_CESS), 2)
                print(f"Estimated Proportionate ITC Reversal calculation done.")
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting estimated_ITC_Reversal: {e}")

            try:
                valuesFrom3b["table_3_1_a_c_e_taxable_value_sum"] = table_3_1_a1 + table_3_1_C1 + table_3_1_E1

                #  Result point 21
                valuesFrom3b["table_4A_row_1_IGST"] = pd.to_numeric(table_4.iloc[1, 1], errors='coerce')
                valuesFrom3b["table_4A_row_1_CGST"] = pd.to_numeric(table_4.iloc[1, 2], errors='coerce')
                valuesFrom3b["table_4A_row_1_SGST"] = pd.to_numeric(table_4.iloc[1, 3], errors='coerce')
                valuesFrom3b["table_4A_row_1_CESS"] = pd.to_numeric(table_4.iloc[1, 4], errors='coerce')
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting Result point 21: {e}")

            #  Result point 20
            try:
                valuesFrom3b["table_4A_row_4_IGST"] = pd.to_numeric(table_4.iloc[4, 1], errors='coerce')
                valuesFrom3b["table_4A_row_4_CGST"] = pd.to_numeric(table_4.iloc[4, 2], errors='coerce')
                valuesFrom3b["table_4A_row_4_SGST"] = pd.to_numeric(table_4.iloc[4, 3], errors='coerce')
                valuesFrom3b["table_4A_row_4_CESS"] = pd.to_numeric(table_4.iloc[4, 4], errors='coerce')
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting Result point 20: {e}")

            # Result point 28 Net ITC available (A-B)
            try:
                table_4c_sum = pd.to_numeric(table_4.iloc[9, 1:], errors='coerce').sum(skipna=True)  # Sum of row 10
                print(f"Table 4C: Net ITC available (A-B): IGST+CGST+SGST+Cess = {table_4c_sum}")
                valuesFrom3b["result_point_28"] = table_4c_sum
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting Result point 28: {e}")

            # 2. Difference of [table 3.1 (d)] -  [ table 4 A (3)] (CGST , SGST , IGST & Cess)
            # table 3.1 sum of row 4, 3rd column onwards : Inward supplies liable to reverse charge -
            # Table 4 A(3) Inward supplies liable to reverse charge (other than 1 & 2 above)
            try:
                valuesFrom3b["diff_In_RCM_ITC_IGST"] = table_3_1.iloc[3, 2] - table_4.iloc[3, 1]
                valuesFrom3b["diff_In_RCM_ITC_CGST"] = table_3_1.iloc[3, 3] - table_4.iloc[3, 2]
                valuesFrom3b["diff_In_RCM_ITC_SGST"] = table_3_1.iloc[3, 4] - table_4.iloc[3, 3]
                valuesFrom3b["diff_In_RCM_ITC_CESS"] = table_3_1.iloc[3, 5] - table_4.iloc[3, 4]
                # print(f"Difference in RCM ITC: {diifIn_RCM_ITC}")
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting diff_In_RCM_ITC: {e}")

            # 3. Calculation : Difference in RCM payment
            # Difference of 3.1(d)] as above -  [table 6.1 B(Total tax payable)]
            # Get last 4 rows from column 1 of Table 6.1 (B) of GSTR-3B_merged.xlsx
            try:
                valuesFrom3b["diff_In_RCM_Pay_IGST"] = table_3_1.iloc[3, 2] - table_6_1.iloc[-4, 1]
                valuesFrom3b["diff_In_RCM_Pay_CGST"] = table_3_1.iloc[3, 3] - table_6_1.iloc[-3, 1]
                valuesFrom3b["diff_In_RCM_Pay_SGST"] = table_3_1.iloc[3, 4] - table_6_1.iloc[-2, 1]
                valuesFrom3b["diff_In_RCM_Pay_CESS"] = table_3_1.iloc[3, 5] - table_6_1.iloc[-1, 1]
                # print(f"Difference in RCM payment done: {diffInRCMPayment}")
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting diff_In_RCM_Pay: {e}")

            # 4. Total tax on Inward supplies (liable to reverse charge) table_3_1_D (CGST + SGST + IGST + Cess)
            try:
                table_3_1_D1_text = "Total tax on" + table_3_1.iloc[3, 0]
                table_3_1_D2 = pd.to_numeric(table_3_1.iloc[3, 2:], errors='coerce').sum(skipna=True)  # Sum of row D
                valuesFrom3b["table_3_1_D1"] = table_3_1_D1_text
                valuesFrom3b["table_3_1_D2"] = table_3_1_D2
                print(f"Total tax on Inward supplies (liable to reverse charge) (CGST + SGST + IGST + Cess): {table_3_1_D2}")
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting Total tax on Inward supplies: {e}")

            # 5. Sum of Table 3.1 a + b + d (CGST + SGST + IGST + Cess)
            try:
                sum_table_3_1_row_a = pd.to_numeric(table_3_1.iloc[0, 2:], errors='coerce').sum(skipna=True)
                sum_table_3_1_row_b = pd.to_numeric(table_3_1.iloc[1, 2:], errors='coerce').sum(skipna=True)
                sum_table_3_1_row_d = pd.to_numeric(table_3_1.iloc[3, 2:], errors='coerce').sum(skipna=True)
                sum_table_3_1_row_a_b_d_taxes = sum_table_3_1_row_a + sum_table_3_1_row_b + sum_table_3_1_row_d
                valuesFrom3b["sum_table_3_1_row_a_b_d_taxes"] = sum_table_3_1_row_a_b_d_taxes
                print(f"Sum 3.1 a + b + d (CGST + SGST + IGST + Cess): {sum_table_3_1_row_a_b_d_taxes}")
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting Sum of Table 3.1 a + b + d: {e}")

            # 6. Table 3.1 (C1 +E1)
            try:
                table_3_1_C1_text = "Taxable value" + table_3_1.iloc[2, 0] + "And" + table_3_1.iloc[4, 0]
                table_3_1_C1_plus_E1 = table_3_1.iloc[2, 1] + table_3_1.iloc[4, 1]
                valuesFrom3b["table_3_1_C1"] = table_3_1_C1_text
                valuesFrom3b["table_3_1_C2"] = table_3_1_C1_plus_E1
                print(f"Table 3.1 (C1 +E1) Taxable value: {table_3_1_C1_plus_E1}")
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting Table 3.1 (C1 +E1): {e}")

            # 7. Table 3.1 (a1+B1+D1+E1-C1)
            try:
                sum_table_3_1_A1_B1_D1_E1_minus_C1 = table_3_1_a1 + table_3_1_B1 + table_3_1_D1 + table_3_1_E1 - table_3_1_C1
                valuesFrom3b["sum_table_3_1_A1_B1_D1_E1_minus_C1"] = sum_table_3_1_A1_B1_D1_E1_minus_C1
                print(f"Table 3.1 (A1+B1+D1+E1-C1) calculation done: {sum_table_3_1_A1_B1_D1_E1_minus_C1}")
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting Table 3.1 (a1+B1+D1+E1-C1): {e}")

            # 8. We need only sum of Total Tax Payable (column 1) of Table 6.1 of GSTR-3B_merged.xlsx
            try:
                total_tax_payable_column_GSTR_3B_table_6 = pd.to_numeric(table_6_1.iloc[:, 1], errors='coerce').sum(
                    skipna=True)
                valuesFrom3b["table_6_1_total_tax_payable_IGST"] = pd.to_numeric(table_6_1.iloc[[1, 6], 1], errors='coerce').sum(
                    skipna=True)
                print(f"valuesFrom3b['table_6_1_total_tax_payable_IGST']: {valuesFrom3b['table_6_1_total_tax_payable_IGST']}")
                valuesFrom3b["table_6_1_total_tax_payable_CGST"] = pd.to_numeric(table_6_1.iloc[[2, 7], 1], errors='coerce').sum(
                    skipna=True)
                print(f"valuesFrom3b['table_6_1_total_tax_payable_CGST']: {valuesFrom3b['table_6_1_total_tax_payable_CGST']}")

                valuesFrom3b["table_6_1_total_tax_payable_SGST"] = pd.to_numeric(table_6_1.iloc[[3, 8], 1], errors='coerce').sum(
                    skipna=True)
                valuesFrom3b["table_6_1_total_tax_payable_CESS"] = pd.to_numeric(table_6_1.iloc[[4, 9], 1], errors='coerce').sum(
                    skipna=True)
                # valuesFrom3b["total_tax_payable_column_GSTR_3B_table_6"] = total_tax_payable_column_GSTR_3B_table_6
            except Exception as e:
                print(f"[GSTR-3B_reader] ❌ Error while setting total_tax_payable Table 6 : {e}")

            #  Parameter 13 of ASMT-10 report
            # In case GSTR-9 is not uploaded, we need to create late fee from GSTR-3B_merged.
            # If table 3.1 (a+b+c+e) total taxable value > Rs 2,00,00,000
            if denominator > 20000000:
                late_fee_calc_from_3B_if_gstr9_not_uploaded = round(0.0025 * denominator, 2)
                valuesFrom3b['result_point_13'] = late_fee_calc_from_3B_if_gstr9_not_uploaded
            else:
                valuesFrom3b['result_point_13'] = 0.0
            print(" === ✅ Returning after execution of file gstr3b_merged_reader.py ===")
            return valuesFrom3b
    except Exception as e:
        print(f"[GSTR-3B_reader] ❌ Error: {e}")
        return valuesFrom3b


def get_stripped_value(cell):
    return cell.strip() if pd.notna(cell) else ""
