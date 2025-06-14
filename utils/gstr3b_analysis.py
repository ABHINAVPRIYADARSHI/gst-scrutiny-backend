from openpyxl import Workbook
from utils.globals.constants import estimatedLiability
from utils.globals.constants import diffInRCMPayment
from utils.globals.constants import diffInRCM_ITC
from .gstr3b_reader import gstr3b_reader
from openpyxl.styles import Alignment

proportionate_reversal = 0

async def generate_gstr3b_analysis(gstin):
    print(" === Starting execution of file gstr3B_analysis.py ===")
    output_path = f"reports/{gstin}/GSTR-3B/GSTR-3B_analysis.xlsx"
    try:
        valuesFrom3b = await gstr3b_reader(gstin)
        # Create workbook and sheet
        wb = Workbook()
        # Define sheet names for each pair
        sheet_names = ["Estimated Liability", "Diif. in RCM ITC", "Diff. in RCM Payment"]
        # Define your data pairs
        data_pairs = [
            ("Estimated Liability", valuesFrom3b[estimatedLiability]),
            ("Difference in RCM ITC", valuesFrom3b[diffInRCM_ITC]),
            ("Difference in RCM payment", valuesFrom3b[diffInRCMPayment]),
        ]
        # Fill each sheet with one row
        for idx, (sheet_name, (label, value)) in enumerate(zip(sheet_names, data_pairs)):
            if idx == 0:
                # Use default sheet for first item and rename it
                ws = wb.active
                ws.title = sheet_name
            else:
                ws = wb.create_sheet(title=sheet_name)

            # Optional: set column width
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 20

            # Write label and rounded value
            ws.cell(row=1, column=1, value=label).alignment = Alignment(wrap_text=True)
            ws.cell(row=1, column=2, value=round(value, 2))

        # Save the workbook
        wb.save(output_path)
        print(f"Excel file saved to: {output_path}")
        print(" === Returning after successful execution 0f file gstr3b_analysis.py ===")

    except Exception as e:
        print(f"[GSTR-3B Analysis] ❌ Error during analysis: {e}")
