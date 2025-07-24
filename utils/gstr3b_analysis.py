import pandas as pd

from utils.globals.constants import result_point_15, result_point_16
from .gstr3b_merged_reader import gstr3b_merged_reader


async def generate_gstr3b_merged_analysis(gstin):
    print(f"[GSTR-3B Analysis] Starting execution of file gstr3B_analysis.py ===")
    final_result_points = {}
    output_path = f"reports/{gstin}/GSTR-3B_analysis.xlsx"
    try:
        valuesFrom3b = await gstr3b_merged_reader(gstin)
        try:
            final_result_points["gstin_of_taxpayer"] = valuesFrom3b["gstin_of_taxpayer"]
            final_result_points["legal_name_of_taxpayer"] = valuesFrom3b["legal_name_of_taxpayer"]
            final_result_points["trade_name_of_taxpayer"] = valuesFrom3b["trade_name_of_taxpayer"]
            final_result_points["financial_year"] = valuesFrom3b["financial_year"]
        except Exception as e:
            print(f"[GSTR-3B Analysis] ❌ Error while setting Taxpayer's info: {e}")

        try:
            final_result_points["result_point_3_IGST"] = valuesFrom3b.get("diff_In_RCM_ITC_IGST")
            final_result_points["result_point_3_CGST"] = valuesFrom3b.get("diff_In_RCM_ITC_CGST")
            final_result_points["result_point_3_SGST"] = valuesFrom3b.get("diff_In_RCM_ITC_SGST")
            final_result_points["result_point_3_CESS"] = valuesFrom3b.get("diff_In_RCM_ITC_CESS")
        except Exception as e:
            print(f"[GSTR-3B Analysis] ❌ Error while setting Taxpayer's info: {e}")

        try:
            final_result_points["result_point_7_IGST"] = valuesFrom3b.get("estimated_ITC_Reversal_IGST")
            final_result_points["result_point_7_CGST"] = valuesFrom3b.get("estimated_ITC_Reversal_CGST")
            final_result_points["result_point_7_SGST"] = valuesFrom3b.get("estimated_ITC_Reversal_SGST")
            final_result_points["result_point_7_CESS"] = valuesFrom3b.get("estimated_ITC_Reversal_CESS")
        except Exception as e:
            print(f"[GSTR-3B Analysis] ❌ Error while setting result_point_7: {e}")

        try:
            final_result_points[result_point_15] = valuesFrom3b.get("table_3_1_a_c_e_taxable_value_sum")
            final_result_points[result_point_16] = valuesFrom3b.get("table_3_1_a_c_e_taxable_value_sum")
        except Exception as e:
            print(f"[GSTR-3B Analysis] ❌ Error while setting result_points 15 & 16: {e}")

        try:
            final_result_points["result_point_20_IGST"] = valuesFrom3b.get("table_4A_row_4_IGST")
            final_result_points["result_point_20_CGST"] = valuesFrom3b.get("table_4A_row_4_CGST")
            final_result_points["result_point_20_SGST"] = valuesFrom3b.get("table_4A_row_4_SGST")
            final_result_points["result_point_20_CESS"] = valuesFrom3b.get("table_4A_row_4_CESS")
        except Exception as e:
            print(f"[GSTR-3B Analysis] ❌ Error while setting result_point_20: {e}")

        try:
            final_result_points["result_point_21_IGST"] = valuesFrom3b.get("table_4A_row_1_IGST")
            final_result_points["result_point_21_CGST"] = valuesFrom3b.get("table_4A_row_1_CGST")
            final_result_points["result_point_21_SGST"] = valuesFrom3b.get("table_4A_row_1_SGST")
            final_result_points["result_point_21_CESS"] = valuesFrom3b.get("table_4A_row_1_CESS")
        except Exception as e:
            print(f"[GSTR-3B Analysis] ❌ Error while setting result_point_21: {e}")

        # Add result point 28 into final map
        try:
            final_result_points["result_point_28"] = valuesFrom3b.get("result_point_28")
        except Exception as e:
            print(f"[GSTR-3B Analysis] ❌ Error while setting result_point_28: {e}")

        # Define sheet names for each pair
        sheet_names = ["Estimated Prop. ITC reversal", "Diif. in RCM ITC", "Diff. in RCM Payment"]
        # Define your data pairs
        try:
            df_est_ITC_reversal = pd.DataFrame([{
                "IGST": valuesFrom3b["estimated_ITC_Reversal_IGST"],
                "CGST": valuesFrom3b["estimated_ITC_Reversal_CGST"],
                "SGST": valuesFrom3b["estimated_ITC_Reversal_SGST"],
                "CESS": valuesFrom3b["estimated_ITC_Reversal_CESS"]
            }])
        except Exception as e:
            print(f"[GSTR-3B Analysis] ❌ Error while computing df_est_ITC_reversal: {e}")

        try:
            df_diff_In_RCM_ITC = pd.DataFrame([{
                "IGST": valuesFrom3b["diff_In_RCM_ITC_IGST"],
                "CGST": valuesFrom3b["diff_In_RCM_ITC_CGST"],
                "SGST": valuesFrom3b["diff_In_RCM_ITC_SGST"],
                "CESS": valuesFrom3b["diff_In_RCM_ITC_CESS"]
            }])
        except Exception as e:
            print(f"[GSTR-3B Analysis] ❌ Error while computing df_diff_In_RCM_ITC: {e}")

        try:
            df_diff_In_RCM_Pay = pd.DataFrame([{
                "IGST": valuesFrom3b["diff_In_RCM_Pay_IGST"],
                "CGST": valuesFrom3b["diff_In_RCM_Pay_CGST"],
                "SGST": valuesFrom3b["diff_In_RCM_Pay_SGST"],
                "CESS": valuesFrom3b["diff_In_RCM_Pay_CESS"]
            }])
        except Exception as e:
            print(f"[GSTR-3B Analysis] ❌ Error while computing df_diff_In_RCM_Pay: {e}")

        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df_est_ITC_reversal.to_excel(writer, sheet_name="Estimated Prop. ITC Reversal", index=False)
            df_diff_In_RCM_ITC.to_excel(writer, sheet_name="Difference in RCM ITC", index=False)
            df_diff_In_RCM_Pay.to_excel(writer, sheet_name="Difference in RCM payment", index=False)

        # Save the workbook
        print(f"Excel file saved to: {output_path}")
        print(" === ✅ Returning after successful execution of file gstr3b_analysis.py ===")
        return final_result_points
    except Exception as e:
        print(f"[GSTR-3B Analysis] ❌ Error during analysis: {e}")
        return final_result_points
