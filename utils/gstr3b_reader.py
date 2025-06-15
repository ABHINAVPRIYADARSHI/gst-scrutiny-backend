import os
import pandas as pd
from utils.globals.constants import OLD_TABLE_POSITIONS_GSTR_3B
from utils.globals.constants import NEW_TABLE_POSITIONS_GSTR_3B
from utils.globals.constants import newFormat
from utils.globals.constants import oldFormat
from utils.globals.constants import estimatedLiability
from utils.globals.constants import diffInRCMPayment
from utils.globals.constants import diffInRCM_ITC
from utils.globals.constants import extract_table_with_header
from utils.globals.constants import convert_to_number


async def gstr3b_reader(gstin):
    print(" === Starting execution of file gstr3B_reader.py ===")
    input_path = f"reports/{gstin}/GSTR-3B/GSTR-3B_merged.xlsx"
    try:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"[GSTR-3B reader]: Input file not found at {input_path}")
        else:
            valuesFrom3b = {}
            # Load full sheet without header
            df_full = pd.read_excel(input_path, sheet_name="GSTR-3B_merged", header=None)
            print("GSTR-3B_merged.xlsx fetched successfully.")
            # Read cell A1 (row 0, column 0)
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

            # 1. Proportionate Reversal/Estimated Liability calculation
            table_3_1 = tables["3.1"].copy()
            table_4 = tables["4"].copy()
            table_6_1 = tables["6.1"].copy()
            # Sum of 6th row (index 5)
            table_4_row6_sum = pd.to_numeric(table_4.iloc[5], errors="coerce").sum(skipna=True)
            A1 = pd.to_numeric(table_3_1.iloc[0, 1], errors='coerce')  # 1st row
            B1 = pd.to_numeric(table_3_1.iloc[1, 1], errors='coerce')  # 2nd row
            C1 = pd.to_numeric(table_3_1.iloc[2, 1], errors='coerce')  # 3rd row
            E1 = pd.to_numeric(table_3_1.iloc[4, 1], errors='coerce')  # 5th row
            numerator = C1 + E1
            denominator = A1 + B1 + C1 + E1
            estimated_liability = round(((numerator / denominator if denominator else 0) * table_4_row6_sum), 2)
            print("Estimated Liability:", estimated_liability)
            valuesFrom3b[estimatedLiability] = estimated_liability
            print("Proportionate Reversal/Estimated Liability calculation done.")

            # 2. Inward supplies liable to reverse charge : sum of row 4, 3rd column onwards of table 3.1
            sum_row_d_table_3_1 = pd.to_numeric(table_3_1.iloc[3, 2:], errors='coerce').sum(skipna=True)
            print(f"Difference in RCM ITC: {sum_row_d_table_3_1}")
            valuesFrom3b[diffInRCM_ITC] = sum_row_d_table_3_1
            print("Inward supplies liable to reverse charge calculation done.")

            # 3. Calculation : Difference in RCM payment
            # Get last 4 rows from column 1 of Table 6.1 of GSTR-3B_merged.xlsx
            sum_table6_column1_last_4_rows = pd.to_numeric(table_6_1.iloc[-4:, 1], errors='coerce').sum(skipna=True)
            print(f"Difference in RCM payment: {sum_table6_column1_last_4_rows}")
            valuesFrom3b[diffInRCMPayment] = sum_table6_column1_last_4_rows
            # We need only sum of Total Tax Payable (column 1) of Table 6.1 of GSTR-3B_merged.xlsx
            total_tax_payable_column_GSTR_3B_table_6 = pd.to_numeric(table_6_1.iloc[:, 1], errors='coerce').sum(
                skipna=True)
            valuesFrom3b["total_tax_payable_column_GSTR_3B_table_6"] = total_tax_payable_column_GSTR_3B_table_6
            print("Calculation : Difference in RCM payment done.")

            # 4. (d) Inward supplies (liable to reverse charge)
            table_3_1_D1 = table_3_1.iloc[3, 0]
            table_3_1_D2 = table_3_1.iloc[3, 1]
            valuesFrom3b["table_3_1_D1"] = table_3_1_D1
            print(f"table_3_1_D1: {table_3_1_D1}")
            valuesFrom3b["table_3_1_D2"] = table_3_1_D2
            print("Inward supplies (liable to reverse charge) calculation done.")

            # 5. Sum 3.1 A1 + C1
            sum_table_3_1_A1_and_C1 = convert_to_number(table_3_1.iloc[0, 1]) + convert_to_number(table_3_1.iloc[2, 1])
            valuesFrom3b["sum_table_3_1_A1_and_C1"] = sum_table_3_1_A1_and_C1
            print("Sum 3.1 A1 + C1 calculation done.")

            # 6. Table 3.1 (C1)
            table_3_1_C1 = table_3_1.iloc[2, 0]
            table_3_1_C2 = table_3_1.iloc[2, 1]
            valuesFrom3b["table_3_1_C1"] = table_3_1_C1
            valuesFrom3b["table_3_1_C2"] = table_3_1_C2
            print("Table 3.1 (C1) calculation done.")

            # 7. Table 3.1 (A1+B1+C1+D1+E1)
            sum_table_3_1_A1_to_E1 = pd.to_numeric(table_3_1.iloc[:, 1], errors='coerce').sum(skipna=True)
            valuesFrom3b["sum_table_3_1_A1_to_E1"] = sum_table_3_1_A1_to_E1
            print("Table 3.1 (A1+B1+C1+D1+E1) calculation done.")
            print(" === ✅ Returning after successful execution 0f file gstr3b_reader.py ===")
            return valuesFrom3b
    except Exception as e:
        print(f"[GSTR-3B_reader] ❌ Error: {e}")
        raise Exception(e)
