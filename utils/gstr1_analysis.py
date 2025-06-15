import pandas as pd
import os

# Constants for fixed column positions
HSN_COL = 0
RATE_COL = 4
TAX_COLS = [5, 6, 7, 8, 9]  # Taxable Value, IGST, CGST, SGST, Cess


async def generate_gstr1_analysis(gstin: str):
    print(f" Started execution of method generate_gstr1_analysis for: {gstin}")
    input_path = f"reports/{gstin}/GSTR-1/GSTR-1_merged.xlsx"
    output_path = f"reports/{gstin}/GSTR-1/GSTR-1_Analysis.xlsx"

    if not os.path.exists(input_path):
        print(f"[GSTR-1 Analysis] Skipped: Input file not found at {input_path}")
        return

    try:
        # Read the HSN sheet with header at second row (index 1)
        df = pd.read_excel(input_path, sheet_name="hsn_merged", header=1)

        # Convert relevant tax columns to numeric
        for col in TAX_COLS:
            df.iloc[:, col] = pd.to_numeric(df.iloc[:, col], errors='coerce')

        # Table 1: Group by HSN
        df_by_hsn = df.groupby(df.iloc[:, HSN_COL])[df.columns[TAX_COLS]].sum().reset_index()

        # Table 2: Group by Rate
        df_by_rate = df.groupby(df.iloc[:, RATE_COL])[df.columns[TAX_COLS]].sum().reset_index()

        # Table 3: Group by HSN and Rate
        # Table 3: Group by HSN and Rate using column positions
        df_by_hsn_rate = df.groupby(
            [df.iloc[:, HSN_COL], df.iloc[:, RATE_COL]]
        )[[df.columns[i] for i in TAX_COLS]].sum().reset_index()
        df_by_hsn_rate.columns = ["HSN", "Rate"] + [df.columns[i] for i in TAX_COLS]

        # Write each table to its own sheet
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_by_hsn.to_excel(writer, index=False, sheet_name="By_HSN")
            df_by_rate.to_excel(writer, index=False, sheet_name="By_Rate")
            df_by_hsn_rate.to_excel(writer, index=False, sheet_name="By_HSN_and_Rate")

        print(f"[GSTR-1 Analysis] ✅ Summary report generated at: {output_path}")
        return output_path

    except Exception as e:
        print(f"[GSTR-1 Analysis] ❌ Error during analysis: {e}")
