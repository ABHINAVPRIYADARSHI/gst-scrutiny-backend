import os
import pandas as pd

# Constants
TAX_COL_NAMES = [
    "Integrated Tax  (₹)",
    "Central Tax (₹)",
    "State/UT tax (₹)",
    "Cess  (₹)"
]
HEADER_ROW = 2
RATE_COL = 8   # Column I (zero-based)
REVERSE_CHARGE_COL =7  # Column H
CANCELLED_DATE_COL = 20  # Column U

def summarize_tax(df: pd.DataFrame, tax_cols):
    """Convert tax columns to numeric and return a one-row summary."""
    df_numeric = df[tax_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
    summary = df_numeric.sum().to_frame().T
    summary["Total Tax"] = summary.sum(axis=1)
    return summary

async def generate_gstr2a_analysis(gstin):
    input_path = f"reports/{gstin}/GSTR-2A/GSTR-2A_merged.xlsx"
    output_path = f"reports/{gstin}/GSTR-2A/GSTR-2A_analysis.xlsx"

    if not os.path.exists(input_path):
        print(f"[GSTR-2A Analysis] Skipped: Input file not found at {input_path}")
        return

    try:
        # Load Excel and extract headers
        df_raw = pd.read_excel(input_path, sheet_name="B2B_merged", header=None)
        df = df_raw.iloc[HEADER_ROW + 1:].reset_index(drop=True)
        df.columns = df_raw.iloc[HEADER_ROW]  # Optional

        # Strip whitespace from all string entries
        df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

        # -- 1. Reverse Charge Summary --
        df_reverse = df[df.iloc[:, REVERSE_CHARGE_COL].astype(str).str.strip() == 'Y']
        summary_reverse = summarize_tax(df_reverse, TAX_COL_NAMES)

        # -- 2. Cancelled ITC Summary --
        df_cancelled = df[df.iloc[:, CANCELLED_DATE_COL].notna()]
        summary_cancelled = summarize_tax(df_cancelled, TAX_COL_NAMES)

        # -- 3. ITC by Rate Summary (without grand total row) --
        df_rate_grouped = (
            df.groupby(df.iloc[:, RATE_COL])[TAX_COL_NAMES]
                .apply(lambda group: group.apply(pd.to_numeric, errors='coerce').fillna(0).sum())
                .reset_index()
                .rename(columns={df.columns[RATE_COL]: "Rate %"})
        )

        # -- Save all summaries to Excel --
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            summary_reverse.to_excel(writer, index=False, sheet_name="Reverse_Charge")
            summary_cancelled.to_excel(writer, index=False, sheet_name="Cancelled_ITC")
            df_rate_grouped.to_excel(writer, index=False, sheet_name="ITC_by_Rate")

        print(f"[GSTR-2A Analysis] ✅ Summary report generated at: {output_path}")

    except Exception as e:
        print(f"[GSTR-2A Analysis] ❌ Error during analysis: {e}")
