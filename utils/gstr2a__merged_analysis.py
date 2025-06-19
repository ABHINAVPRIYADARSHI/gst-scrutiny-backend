import os
import pandas as pd
from tabulate import tabulate
from utils.globals.constants import result_point_5, result_point_14, result_point_19, result_point_20

TAX_COL_NAMES = [
    "Integrated Tax  (₹)",
    "Central Tax (₹)",
    "State/UT tax (₹)",
    "Cess  (₹)"
]
ECOM_COL_NAMES = [8,9,10,11,12]
TCS_COL_NAMES = [5,6,7,8]  # Cess is absent in both TDS & TCS merged tables
TDS_COL_NAMES = [3,4,5,6]
HEADER_ROW = 2
RATE_COL = 8   # Column I (zero-based)
REVERSE_CHARGE_COL =7  # Column H
CDNR_REVERSE_CHARGE_COL =8  # Column I
CDNR_NOTE_TYPE_COL = 2  # Column C
CANCELLED_DATE_COL = 20  # Column U
sheet_names = ["B2B_merged", "ISD_merged", "IMPG_merged", "IMPG SEZ_merged"]


async def generate_gstr2a_merged_analysis(gstin):
    print(f" === Started execution of method generate_gstr2a_analysis for: {gstin} ===")
    final_result_points = {}
    input_path = f"reports/{gstin}/GSTR-2A_merged.xlsx"
    output_path = f"reports/{gstin}/GSTR-2A_analysis.xlsx"

    if not os.path.exists(input_path):
        print(f"[GSTR-2A Analysis] Skipped: Input file not found at {input_path}")
        return final_result_points

    try:
        # Load Excel and extract headers
        df_raw = pd.read_excel(input_path, sheet_name="B2B_merged", header=None)
        df_B2B_merged = df_raw.iloc[HEADER_ROW + 1:].reset_index(drop=True)
        df_B2B_merged.columns = df_raw.iloc[HEADER_ROW]  # Optional

        # Strip whitespace from all string entries
        df_B2B_merged = df_B2B_merged.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

        # -- 1. B2B_merged sheet: Reverse Charge Summary --
        df_reverse_B2B = df_B2B_merged[df_B2B_merged.iloc[:, REVERSE_CHARGE_COL].astype(str).str.strip() == 'Y']
        reverse_charge_liability_B2B_merged = summarize_tax(df_reverse_B2B, TAX_COL_NAMES)
        print("Reverse Charge liability calculation completed.")

        # -- 2. B2B_merged sheet: Cancelled ITC Summary (CANCELLED_DATE_COL should not be empty)--
        itc_from_cancelled_taxpayers_B2B = df_B2B_merged[df_B2B_merged.iloc[:, CANCELLED_DATE_COL].notna()]
        itc_from_cancelled_taxpayers_B2B_merged = summarize_tax(itc_from_cancelled_taxpayers_B2B, TAX_COL_NAMES)
        final_result_points[result_point_5] = itc_from_cancelled_taxpayers_B2B_merged.iloc[0, -1]
        print("Cancelled ITC Summary calculation completed.")

        # -- 3. ITC by Rate Summary (without grand total row) --
        df_rate_grouped = (
            df_B2B_merged.groupby(df_B2B_merged.iloc[:, RATE_COL])[TAX_COL_NAMES]
                .apply(lambda group: group.apply(pd.to_numeric, errors='coerce').fillna(0).sum())
                .reset_index()
                .rename(columns={df_B2B_merged.columns[RATE_COL]: "Rate %"})
        )
        print("ITC by Rate Summary calculation completed.")

        # -- 4. Total ITC as per 2A (take data rows from sheets B2B_merged + ISD_merged + IMPG_merged +
        #       IMPG_SEZ_merged + Net of CDNR_merged where reverse_charge column = 'N' ). We assume that
        #       3rd row is the header and contains columns "Integrated Tax  (₹)", "Central Tax (₹)",
        #       "State/UT tax (₹)", "Cess (₹)"
        print(f"Sheet Total ITC as per 2A calculation started.")
        summary_rows = []
        for sheet in sheet_names:
            print(f"Evaluating sheet: {sheet}")
            df_raw = pd.read_excel(input_path, sheet_name=sheet, header=None)
            df = df_raw.iloc[HEADER_ROW + 1:].reset_index(drop=True)
            df.columns = df_raw.iloc[HEADER_ROW]
            # Strip whitespace from all string entries
            df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))
            if sheet == "B2B_merged":
                df = df[df.iloc[:, REVERSE_CHARGE_COL].astype(str).str.strip() == 'N']
            if not df.empty:
                tax_sum = []
                for col in TAX_COL_NAMES:
                    if col in df.columns:
                        col_sum = pd.to_numeric(df[col], errors="coerce").sum()
                    else:
                        col_sum = 0
                    tax_sum.append(col_sum)
            else:
                tax_sum = [0] * len(TAX_COL_NAMES)
            # Optionally include the sheet name in the summary row
            summary_rows.append([sheet] + tax_sum)
        # CDNR = Credit Debit Note Regular
        print(f"Evaluating sheet: CDNR_merged")
        cdnr_raw = pd.read_excel(input_path, sheet_name="CDNR_merged", header=None)
        cdnr_df = cdnr_raw.iloc[HEADER_ROW + 1:].reset_index(drop=True)
        cdnr_df.columns = cdnr_raw.iloc[HEADER_ROW]
        cdnr_df = cdnr_df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))
        cdnr_df = cdnr_df[cdnr_df.iloc[:, CDNR_REVERSE_CHARGE_COL].astype(str).str.strip() == 'N']
        # First append debit note values. Later on append credit note values. Ensure that credit
        # note is the last row to be appended because that is subtracted.
        debit_df = cdnr_df[cdnr_df.iloc[:, CDNR_NOTE_TYPE_COL].astype(str).str.strip() == 'D']
        if not debit_df.empty:
            tax_sum = debit_df[TAX_COL_NAMES].apply(pd.to_numeric, errors="coerce").sum().tolist()
        else:
            tax_sum = [0] * len(TAX_COL_NAMES)
        summary_rows.append(["Debit Note"] + tax_sum)
        credit_df = cdnr_df[cdnr_df.iloc[:, CDNR_NOTE_TYPE_COL].astype(str).str.strip() == 'C']
        if not credit_df.empty:
            tax_sum = credit_df[TAX_COL_NAMES].apply(pd.to_numeric, errors="coerce").sum().tolist()
        else:
            tax_sum = [0] * len(TAX_COL_NAMES)
        summary_rows.append(["Credit Note"] + tax_sum)
        summary_df = pd.DataFrame(summary_rows, columns=["Sheet Name", TAX_COL_NAMES[0], TAX_COL_NAMES[1],
                                                         TAX_COL_NAMES[2], TAX_COL_NAMES[3]])
        isd_total = summary_df.loc[summary_df["Sheet Name"] == "ISD_merged", TAX_COL_NAMES].astype(float).sum(axis=1).values[0]
        impg_total = summary_df.loc[summary_df["Sheet Name"] == "IMPG_merged", TAX_COL_NAMES].astype(float).sum(axis=1).values[0]
        impg_sez_total = summary_df.loc[summary_df["Sheet Name"] == "IMPG SEZ_merged", TAX_COL_NAMES].astype(float).sum(axis=1).values[0]
        impg_and_impg_sez_total = impg_total + impg_sez_total
        final_result_points[result_point_19] = isd_total
        final_result_points[result_point_20] = impg_and_impg_sez_total
        # debit note values are added, credit note values are subtracted
        # Separate credit note values (last row)
        credit_note_values = summary_df.iloc[-1][TAX_COL_NAMES].apply(pd.to_numeric, errors="coerce")
        # Sum all rows except the last one
        all_except_credit = summary_df.iloc[:-1][TAX_COL_NAMES].apply(pd.to_numeric, errors="coerce").sum()
        # Compute final totals: sum(including debit note) - credit note
        final_total = all_except_credit - credit_note_values
        # Append Total row
        summary_df.loc[len(summary_df.index)] = ["Total"] + final_total.tolist()
        print("Total ITC as per 2A calculation completed.")

        # 5. Take full sheet ECOM_merged
        ecom_df_raw = pd.read_excel(input_path, sheet_name="ECO_merged", header=None)
        ecom_df = ecom_df_raw.iloc[HEADER_ROW + 1:].reset_index(drop=True)
        ecom_df.columns = ecom_df_raw.iloc[HEADER_ROW]
        # Strip whitespace from all string entries
        ecom_df = ecom_df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))
        ecom_df = ecom_df.iloc[:, ECOM_COL_NAMES]
        ecom_df.columns.values[0] = "Taxable Value"
        ecom_df_total_row = []
        for i, col in enumerate(ecom_df.columns):
            if i == 0:
                ecom_df_total_row.append('TOTAL')  # First column = label
            else:
                ecom_df_total_row.append(ecom_df[col].astype(float).sum())  # Sum each column
        ecom_df_total_row = pd.DataFrame([ecom_df_total_row], columns=ecom_df.columns)  # Append total row
        ecom_df = pd.concat([ecom_df, ecom_df_total_row], ignore_index=True)
        print("ECOM_merged calculation completed.")

        # 6. Take full sheet TCS_merged
        tcs_df_raw = pd.read_excel(input_path, sheet_name="TCS_merged", header=None)
        tcs_df = tcs_df_raw.iloc[HEADER_ROW + 1:].reset_index(drop=True)
        tcs_df.columns = tcs_df_raw.iloc[HEADER_ROW]
        # Strip whitespace from all string entries
        tcs_df = tcs_df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))
        tcs_df = tcs_df.iloc[ :, TCS_COL_NAMES]
        tcs_df.columns.values[0] = "Taxable Value"
        tcs_df.insert(0, "Label", "")
        tcs_df_total_row = ["TOTAL"]
        for col in tcs_df.columns[1:]:
            try:
                total = tcs_df[col].astype(float).sum()
            except ValueError:
                total = ""
            tcs_df_total_row.append(total)
        tcs_df_total_row = pd.DataFrame([tcs_df_total_row], columns=tcs_df.columns)
        tcs_df = pd.concat([tcs_df, tcs_df_total_row], ignore_index=True)
        tcs_taxable_value_total = tcs_df["Taxable Value"].iloc[-1]
        final_result_points[result_point_14] = tcs_taxable_value_total

        print("TCS_merged calculation completed.")

        # 7. Take full sheet TDS_merged
        tds_df_raw = pd.read_excel(input_path, sheet_name="TDS_merged", header=None)
        tds_df = tds_df_raw.iloc[HEADER_ROW + 1:].reset_index(drop=True)
        tds_df.columns = tds_df_raw.iloc[HEADER_ROW]
        # Strip whitespace from all string entries
        tds_df = tds_df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))
        tds_df = tds_df.iloc[ :, TDS_COL_NAMES]
        tds_df.columns.values[0] = "Taxable Value"
        tds_df.insert(0, "Label", "")
        tds_df_total_row = ["TOTAL"]
        for col in tds_df.columns[1:]:  # Skip 'Label' column
            try:
                total = tds_df[col].astype(float).sum()
            except ValueError:
                total = ""  # Leave blank if not numeric
            tds_df_total_row.append(total)
        tds_df_total_row = pd.DataFrame([tds_df_total_row], columns=tds_df.columns)
        tds_df = pd.concat([tds_df, tds_df_total_row], ignore_index=True)
        tds_taxable_value_total = tds_df["Taxable Value"].iloc[-1]
        final_result_points[result_point_14] = tds_taxable_value_total
        print("TDS_merged calculation completed.")

        # -- Save all summaries to Excel --
        print("Started writing all dataframes to GSTR-2A_analysis.xlsx.")
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            reverse_charge_liability_B2B_merged.to_excel(writer, index=False, sheet_name="Reverse_Charge_Liability")
            itc_from_cancelled_taxpayers_B2B_merged.to_excel(writer, index=False, sheet_name="Cancelled_ITC")
            df_rate_grouped.to_excel(writer, index=False, sheet_name="ITC_by_Rate")
            summary_df.to_excel(writer, index= False, sheet_name="Total ITC as per 2A")
            ecom_df.to_excel(writer, index= False, sheet_name="Total ECOM merged")
            tds_df.to_excel(writer, index= False, sheet_name="Total TDS merged")
            tcs_df.to_excel(writer, index= False, sheet_name="Total TCS merged")
        print(f"[GSTR-2A Analysis] ✅ Summary report generated at: {output_path}")
        return final_result_points
    except Exception as e:
        print(f"[GSTR-2A Analysis] ❌ Error during analysis: {e}")
        return final_result_points


def summarize_tax(df: pd.DataFrame, tax_cols):
    """Convert tax columns to numeric and return a one-row summary."""
    df_numeric = df[tax_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
    summary = df_numeric.sum().to_frame().T
    summary["Total Tax"] = summary.sum(axis=1)
    return summary