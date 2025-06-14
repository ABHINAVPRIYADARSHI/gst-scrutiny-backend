import os
from .gstr1_merged import generate_gstr1_merged
from .gstr1_analysis import generate_gstr1_analysis
from .gstr2a_merged import generate_gstr2a_merged
from .gstr2a_analysis import generate_gstr2a_analysis
from .gstr3b_merged import generate_gstr3b_merged
from .gstr3b_analysis import generate_gstr3b_analysis
from .gstr9_analysis import generate_gstr9_analysis
from .ewb_in_merged import generate_ewb_in_merged
from .ewb_out_merged import generate_ewb_out_merged

return_types = ["GSTR-1", "GSTR-2A", "GSTR-3B", "GSTR-9", "EWB-IN", "EWB-OUT"]  # Add more return types when needed
generated_reports = []


async def generate_merged_excel_and_analysis_report(gstin):
    await generate_merged_excel_for_return_types(gstin)
    await generate_analysis_reports(gstin)


async def generate_analysis_reports(gstin):
    await generate_gstr1_analysis(gstin)
    await generate_gstr2a_analysis(gstin)
    await generate_gstr3b_analysis(gstin)
    await generate_gstr9_analysis(gstin)


async def generate_merged_excel_for_return_types(gstin):
    generated_reports = []
    for rt in return_types:
        input_dir = f"uploaded_files/{gstin}/{rt}"
        output_dir = f"reports/{gstin}/{rt}"
        os.makedirs(output_dir, exist_ok=True)

        if not os.path.exists(input_dir) or not os.listdir(input_dir):
            print(f"[{rt}] merge Skipped: No input files in {input_dir}")
            continue

        try:
            match rt:
                case "GSTR-1":
                    output_file = await generate_gstr1_merged(input_dir, output_dir)
                    generated_reports.append({"return_type": rt, "report": output_file})
                    # case "GSTR-2A":
                    await generate_gstr2a_merged(input_dir, output_dir)
                    generated_reports.append({"return_type": rt, "report": output_file})
                case "GSTR-3B":
                    output_file = await generate_gstr3b_merged(input_dir, output_dir)
                    generated_reports.append({"return_type": rt, "report": output_file})
                # case "GSTR-9":
                #      await generate_gstr9(gstin, input_dir, output_dir)
                #      generated_reports.append({"return_type": rt, "report": output_file})
                case "EWB-IN":
                    await generate_ewb_in_merged(input_dir, output_dir)
                    generated_reports.append({"return_type": rt, "report": output_file})
                case "EWB-OUT":
                    await generate_ewb_out_merged(input_dir, output_dir)
                    generated_reports.append({"return_type": rt, "report": output_file})
                case _:
                    print(f" Not a valid return type  {rt}")

        except Exception as e:
            print(f"[{rt}] Error: {e}")
            continue
