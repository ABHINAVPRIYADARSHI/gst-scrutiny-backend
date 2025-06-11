import os
import pandas as pd

# Constants: start and end positions of tables (0-indexed)
TABLE_POSITIONS = {
    "3.1": {
        "start_row": 1,  # Excel row 2
        "end_row": 6,    # Excel row 7
        "start_col": 0,   # Column A
        "end_col": 5      # Column F
    },
    "3.1.1": {
        "start_row": 9,  # Excel row 10
        "end_row": 11,  # Excel row 12
        "start_col": 0,  # Column A
        "end_col": 5  # Column F
    },
    "3.2": {
        "start_row": 14,  # Excel row 15
        "end_row": 17,  # Excel row 18
        "start_col": 0,  # Column A
        "end_col": 2  # Column C
    },
    "4": {
        "start_row": 20,  # Excel row 21
        "end_row": 33,    # Excel row 34
        "start_col": 0, # Column A
        "end_col": 4 # Column E
    },
    "5": {
        "start_row": 36,  # Excel row 37
        "end_row": 38,  # Excel row 39
        "start_col": 0,  # Column A
        "end_col": 2  # Column C
    },
    "5.1": {
        "start_row": 41,  # Excel row 42
        "end_row": 44,  # Excel row 45
        "start_col": 0,  # Column A
        "end_col": 4  # Column E
    },
    "6.1": {
        "start_row": 48,  # Excel row 47
        "end_row": 58,  # Excel row 57
        "start_col": 0,  # Column A
        "end_col": 8  # Column I
    },
}
proportionate_reversal = 0

async def generate_gstr3b_analysis(gstin):
    input_path = f"reports/{gstin}/GSTR-3B/GSTR-3B_merged.xlsx"
    output_path = f"reports/{gstin}/GSTR-3B/GSTR-3B_analysis.xlsx"

    if not os.path.exists(input_path):
        print(f"[GSTR-3B Analysis] Skipped: Input file not found at {input_path}")
        return
    try:
        # Load full sheet without header
        df_full = pd.read_excel(input_path, sheet_name="GSTR-3B_merged", header=None)
        # Extract all tables
        tables = {}
        for key in TABLE_POSITIONS:
            tables[key] = extract_table_with_header(df_full, key, TABLE_POSITIONS)

        # 1. Proportionate Reversal calculation
        table_3_1 = tables["3.1"].copy()
        table_4 = tables["4"].copy()
        table_3_1 = table_3_1.apply(pd.to_numeric, errors='ignore')
        table_4 = table_4.apply(pd.to_numeric, errors='coerce')
        # Sum of 6th row (index 5)
        table_4_row6_sum = table_4.iloc[5].sum(skipna=True)
        A1 = pd.to_numeric(table_3_1.iloc[0, 1], errors='coerce')  # 1st row
        B1 = pd.to_numeric(table_3_1.iloc[1, 1], errors='coerce')  # 2nd row
        C1 = pd.to_numeric(table_3_1.iloc[2, 1], errors='coerce')  # 3rd row
        E1 = pd.to_numeric(table_3_1.iloc[4, 1], errors='coerce')  # 5th row
        numerator = C1 + E1
        denominator = A1 + B1 + C1 + E1
        print(f"Numerator: {numerator}")
        print(f"Denominator: {denominator}")
        print(f"table_4_row5_sum: {table_4_row6_sum}")

        proportionate_reversal = (numerator / denominator if denominator else 0) * table_4_row6_sum
        print("Proportionate Reversal:", proportionate_reversal)

        # #2 Inward supplies liable to reverse charge : sum of row 4, 3rd column onwards of table 3.1
        row_d_sum = pd.to_numeric(table_3_1.iloc[3, 2:], errors='coerce').sum(skipna=True)
        print(f"row_d_sum: {row_d_sum}")

        # # 3. Calculation 3
        table_6_1 = tables["6.1"].copy().apply(pd.to_numeric, errors='coerce')
        # # Get last 4 rows from column index 1
        last_4_values = (table_6_1.iloc[-4:, 1]).sum(skipna=True)
        print(last_4_values)
        # result_df = pd.DataFrame([[last_4_sum]], columns=["Reverse charge paid in cash (Table 6.1)"])

        # --- Minimal output ---
        result_df = pd.DataFrame([
            ["Proportionate Reversal", proportionate_reversal]
        ], columns=["Description", "Value"])

        # --- Export ---
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            result_df.to_excel(writer, sheet_name="Proportionate_Reversal", index=False)

        print("✅ Written: Proportionate Reversal to GSTR-3B_analysis.xlsx")
    except Exception as e:
        print(f"[GSTR-3B Analysis] ❌ Error during analysis: {e}")

def extract_table_with_header(df, table_key, table_positions):
    pos = table_positions[table_key]
    raw_table = df.iloc[
        pos["start_row"]: pos["end_row"] + 1,
        pos["start_col"]: pos["end_col"] + 1
    ]
    # Use first row as header
    raw_table.columns = raw_table.iloc[0]
    cleaned_table = raw_table.drop(raw_table.index[0]).reset_index(drop=True)
    return cleaned_table
