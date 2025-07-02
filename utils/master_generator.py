import os

from .asmt_report_generator import asmt_10_report_generator
from .general_report_generator import general_analysis_report_generator
from .bo_comparison_summary_analysis import generate_bo_comparison_summary_analysis
from .ewb_in_merged import generate_ewb_in_merged
from .ewb_in_merged_analysis import generate_ewb_in_merged_analysis
from .ewb_out_merged import generate_ewb_out_merged
from .ewb_out_merged_analysis import generate_ewb_out_merged_analysis
from .gstr1_merged import generate_gstr1_merged
from .gstr1_merged_analysis import generate_gstr1_merged_analysis
from .gstr2a__merged_analysis import generate_gstr2a_merged_analysis
from .gstr2a_merged import generate_gstr2a_merged
from .gstr3b_analysis import generate_gstr3b_merged_analysis
from .gstr3b_merged_writer import generate_gstr3b_merged
from .gstr9_Vs_3B_analysis import generate_gstr9_Vs_3B_analysis

return_types = ["GSTR-1", "GSTR-2A", "GSTR-3B", "EWB-IN", "EWB-OUT"]


async def generate_merged_excel_and_analysis_report(gstin, report_flag):
    master_dict = {'details_of_taxpayer': {'gstin_of_taxpayer': gstin}}  # Analysis points gathered to populate ASTM-10 sheet
    generated_reports = []   # List of merged files generated
    generated_reports = await generate_merged_excel_for_return_types(gstin, generated_reports, master_dict)
    print(f"[Master Generator] Starting with analysis of merged reports for GSTIN: {gstin}")
    await generate_return_type_reports(gstin, master_dict)
    await general_analysis_report_generator(gstin, master_dict)
    if report_flag:
        await asmt_10_report_generator(gstin, master_dict)
    print(f"generated_reports: {generated_reports}")
    return generated_reports


async def generate_return_type_reports(gstin, master_dict):
    gstr1_analysis_dict = await generate_gstr1_merged_analysis(gstin)
    gstr2a_analysis_dict = await generate_gstr2a_merged_analysis(gstin)
    gstr3b_analysis_dict = await generate_gstr3b_merged_analysis(gstin)
    gstr9_Vs_3b_analysis_dict = await generate_gstr9_Vs_3B_analysis(gstin)
    ewb_in_analysis_dict = await generate_ewb_in_merged_analysis(gstin)
    ewb_out_analysis_dict = await generate_ewb_out_merged_analysis(gstin)
    bo_comparison_summary_dict = await generate_bo_comparison_summary_analysis(gstin)

    if gstr1_analysis_dict:
        master_dict["gstr1_analysis_dict"] = gstr1_analysis_dict
    if gstr2a_analysis_dict:
        master_dict["gstr2a_analysis_dict"] = gstr2a_analysis_dict
    if gstr3b_analysis_dict:
        master_dict["gstr3b_analysis_dict"] = gstr3b_analysis_dict
    if gstr9_Vs_3b_analysis_dict:
        master_dict["gstr9_Vs_3b_analysis_dict"] = gstr9_Vs_3b_analysis_dict
    # if ewb_in_analysis_dict:
    #     master_dict["ewb_in_analysis_dict"] = ewb_in_analysis_dict
    if ewb_out_analysis_dict:
        master_dict["ewb_out_analysis_dict"] = ewb_out_analysis_dict
    if bo_comparison_summary_dict:
        master_dict["bo_comparison_summary_dict"] = bo_comparison_summary_dict


async def generate_merged_excel_for_return_types(gstin, generated_reports, master_dict):
    print(f"[Master Generator] Starting execution of function generate_merged_excel_for_return_types for GSTIN: {gstin} ===")
    for rt in return_types:
        input_dir = f"uploaded_files/{gstin}/{rt}"
        output_dir = f"reports/{gstin}/"
        os.makedirs(output_dir, exist_ok=True)

        if not os.path.exists(input_dir) or not os.listdir(input_dir):
            print(f"[{rt}] merge Skipped: No input files in {input_dir}")
            continue

        try:
            match rt:
                case "GSTR-1":
                    output_file, dict_1_merged = await generate_gstr1_merged(input_dir, output_dir)
                    generated_reports.append({"return_type": rt, "report": output_file}) if output_file else None
                    master_dict["gstr1_merged_dict"] = dict_1_merged
                case "GSTR-2A":
                    output_file = await generate_gstr2a_merged(input_dir, output_dir)
                    generated_reports.append({"return_type": rt, "report": output_file}) if output_file else None
                case "GSTR-3B":
                    output_file, dict_3b_merged = await generate_gstr3b_merged(input_dir, output_dir)
                    generated_reports.append({"return_type": rt, "report": output_file}) if output_file else None
                    master_dict["gstr3b_merged_dict"] = dict_3b_merged
                # case "GSTR-9":
                #      We don't merge GSTR-9, we directly analyse it as its a single file.
                case "EWB-IN":
                    output_file = await generate_ewb_in_merged(input_dir, output_dir)
                    generated_reports.append({"return_type": rt, "report": output_file}) if output_file else None
                case "EWB-OUT":
                    output_file = await generate_ewb_out_merged(input_dir, output_dir)
                    generated_reports.append({"return_type": rt, "report": output_file}) if output_file else None
                case _:
                    print(f" Not a valid return type  {rt}")
        except Exception as e:
            print(f"[{rt}] Error: {e}")
            continue

    print(f"✅ Function call generate_merged_excel_for_return_types completed for GSTIN: {gstin} ===")
    return generated_reports



