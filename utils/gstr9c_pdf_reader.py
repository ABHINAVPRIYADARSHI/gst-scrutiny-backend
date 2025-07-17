import os
import pandas as pd
from glob import glob
import pdfplumber
import datetime
from tabulate import tabulate
from utils.globals.constants import int_eighteen, clean_and_parse_number, newFormat, oldFormat, int_twenty_one

financial_year_2019_20 = "2019-20"
financial_year_2020_21 = "2020-21"

# Define how many rows to skip per table before setting headers. Don't comment out any line else uploaded
# table_position_in_useful_tables will have to be changed accordingly. Out of all tables extracted from the
# PDF file, only the tables whose positions are mentioned as keys in table_header_rows_skip will be saved in
# useful_tables. The table name : position as per GSTR-9C PDF file is stored in table_position_in_useful_tables.
table_header_rows_skip_gstr9c_old_format = {
    0: -1,  # We are not dropping any row
    2: -1,  # Table 5 Part II
    8: -1,  # Table 9 Part II
    10: 2,  # Table 11 Part I
    11: -1,  # Table 11 Part II
    12: 1,  # Table 12
    18: 0,  # Table 16
}
table_header_rows_skip_gstr9c_new_format = {
    0: -1,
    2: -1,  # Table 5 Part II
    8: -1,  # Table 9 Part II
    10: 2,  # Table 11 Part I
    11: -1,  # Table 11 Part II
    12: 1,  # Table 12
    17: 0,  # Table 16
}
header_row_map_old_format = {
    0: 0,
    10: 2,
    12: 2,
    18: 1
}
header_row_map_new_format = {
    0: 0,
    1: 2,
    12: 2,
    17: 1
}
table_position_in_useful_tables_gstr9c = {
    "Table_1": 0,
    "Table_5_part_II": 1,
    "Table_9_part_II": 2,
    "Table_11_part_I": 3,
    "Table_11_part_II": 4,
    "Table_12": 5,
    "Table_16": 6,
}


async def gstr9c_pdf_reader(gstin):
    print(f"[GSTR-9C reader] Starting execution of file gstr9c_pdf_reader.py ===")
    all_tables = []
    useful_tables = []
    valuesFrom9c = {}
    gstr9c_format = oldFormat  # Let by-default be OLD_FORMAT
    input_path_of_GSTR_9 = f"uploaded_files/{gstin}/GSTR-9C/"
    output_path_GSTR_9 = f"reports/{gstin}/GSTR-9C.xlsx"

    try:
        pdf_files = glob(os.path.join(input_path_of_GSTR_9, "*.pdf"))
        if not pdf_files:
            raise FileNotFoundError(f"[GSTR-9reader]: Input file not found at {input_path_of_GSTR_9}")
        print(f"Found {len(pdf_files)} GSTR-9 PDF file(s).")

        # Read the only annual GSTR-9 file and extract tables
        with pdfplumber.open(pdf_files[0]) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    all_tables.append(table)
        print(f" No. of tables in GSTR-9 PDF: {len(all_tables)}")

        if len(all_tables) == int_twenty_one:  # Old format = 21, New format =  tables
            table_header_rows_skip = table_header_rows_skip_gstr9c_old_format
            header_row_map = header_row_map_old_format
            print(f"GSTR-9C.pdf is based on old format")
        else:
            gstr9c_format = newFormat
            table_header_rows_skip = table_header_rows_skip_gstr9c_new_format
            header_row_map = header_row_map_new_format
            print(f"GSTR-9C.pdf is based on new format")

        # Clean and process specific tables of GSTR-9
        for idx, skip_rows in table_header_rows_skip.items():
            if idx < len(all_tables):
                df = pd.DataFrame(all_tables[idx])
                if idx in header_row_map:
                    df.columns = df.iloc[header_row_map.get(idx)]  # Set new header
                df = df.drop(index=range(skip_rows + 1))  # Drop the unnecessary rows
                df = df.reset_index(drop=True)  # Reset index
                useful_tables.append(df)
                # Few tables can't be printed using tabulate due to its merged cell structure
                # print(f"Table no: {idx}")
                # print(tabulate(df, tablefmt='grid', maxcolwidths=20))
        print(f"[GSTR-9C] useful_tables size: {len(useful_tables)}")  # It should be 7 for both old & new.

        for idx, df in enumerate(useful_tables):
            if idx == 0:
                continue
            for row_idx in range(df.shape[0]):
                for col_idx in range(2, df.shape[1]):
                    df.iat[row_idx, col_idx] = clean_and_parse_number(df.iat[row_idx, col_idx])

        # Write the dataframes in excel sheet GSTR-9.xlsx
        with pd.ExcelWriter(output_path_GSTR_9, engine="xlsxwriter") as writer:
            for i, df in enumerate(useful_tables):
                sheet_name = f"Table_{i}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]
                # Create a text wrap format
                workbook = writer.book
                wrap_format = workbook.add_format({'text_wrap': True})
                # Apply wrap format + set column width 20 for all columns
                worksheet.set_column(0, len(df.columns) - 1, 20, wrap_format)
            print("GSTR-9c.xlsx file written successfully...")

        # Para 23: Table 5R - unreconciled turnover value- return value of +ve else zero.
        try:
            table_5_part_II = useful_tables[table_position_in_useful_tables_gstr9c["Table_5_part_II"]]
            # table_5_R = table_5_part_II.iloc[len(table_5_part_II) - 1, 3]  # Row R is last row
            row_R = table_5_part_II[table_5_part_II.iloc[:, 0] == 'R']
            valuesFrom9c["table_5_R"] = row_R.iloc[0, 3]
            print(f"[GSTR-9C] Table 5R - unreconciled turnover value: {valuesFrom9c['table_5_R']}")
        except Exception as e:
            print(f"[GSTR-9C reader] Error while computing Table5 R: {e}")

        # Para 24: Table 9R - Unreconciled tax payment - Return value if -ve else return zero.
        try:
            table_9_part_II = useful_tables[table_position_in_useful_tables_gstr9c["Table_9_part_II"]]
            row_R = table_9_part_II[table_9_part_II.iloc[:, 0] == 'R']
            # table_9_R_CGST = table_9_part_II.iloc[len(table_9_part_II) - 1, 3]  # Row R is last row
            # table_9_R_SGST = table_9_part_II.iloc[len(table_9_part_II) - 1, 4]  # Row R is last row
            # table_9_R_IGST = table_9_part_II.iloc[len(table_9_part_II) - 1, 5]  # Row R is last row
            # table_9_R_CESS = table_9_part_II.iloc[len(table_9_part_II) - 1, 6]  # Row R is last row
            table_9_R_CGST = row_R.iloc[0, 3]
            table_9_R_SGST = row_R.iloc[0, 4]
            table_9_R_IGST = row_R.iloc[0, 5]
            table_9_R_CESS = row_R.iloc[0, 6]
            print(f"[GSTR-9C] Table 9R - Unreconciled tax payment CGST: {table_9_R_CGST}")
            print(f"[GSTR-9C] Table 9R - Unreconciled tax payment SGST: {table_9_R_SGST}")
            print(f"[GSTR-9C] Table 9R - Unreconciled tax payment IGST: {table_9_R_IGST}")
            print(f"[GSTR-9C] Table 9R - Unreconciled tax payment CESS: {table_9_R_CESS}")
            valuesFrom9c["table_9_R_CGST"] = table_9_R_CGST
            valuesFrom9c["table_9_R_SGST"] = table_9_R_SGST
            valuesFrom9c["table_9_R_IGST"] = table_9_R_IGST
            valuesFrom9c["table_9_R_CESS"] = table_9_R_CESS
        except Exception as e:
            print(f"[GSTR-9C reader] Error while computing Table9 R: {e}")

        # Para 25: Table 11- Additional amount payable in Cash - CGST total, SGST total,
        # IGST total, Cess total
        # Merge tables 11 Part I & Part II
        try:
            table_11_part_I = useful_tables[table_position_in_useful_tables_gstr9c["Table_11_part_I"]]
            table_11_part_II = useful_tables[table_position_in_useful_tables_gstr9c["Table_11_part_II"]]
            table_11_part_II.columns = table_11_part_I.columns
            table_11_merged = pd.concat([table_11_part_I, table_11_part_II], ignore_index=True)
            table_11_CGST_total = pd.to_numeric(table_11_merged.iloc[:, 3], errors='coerce').sum(skipna=True)
            table_11_SGST_total = pd.to_numeric(table_11_merged.iloc[:, 4], errors='coerce').sum(skipna=True)
            table_11_IGST_total = pd.to_numeric(table_11_merged.iloc[:, 5], errors='coerce').sum(skipna=True)
            table_11_CESS_total = pd.to_numeric(table_11_merged.iloc[:, 6], errors='coerce').sum(skipna=True)
            print(f"[GSTR-9C] Table 11 CGST total: {table_11_CGST_total}")
            print(f"[GSTR-9C] Table 11 SGST total: {table_11_SGST_total}")
            print(f"[GSTR-9C] Table 11 IGST total: {table_11_IGST_total}")
            print(f"[GSTR-9C] Table 11 CESS total: {table_11_CESS_total}")
            valuesFrom9c["table_11_CGST_total"] = table_11_CGST_total
            valuesFrom9c["table_11_SGST_total"] = table_11_SGST_total
            valuesFrom9c["table_11_IGST_total"] = table_11_IGST_total
            valuesFrom9c["table_11_CESS_total"] = table_11_CESS_total
        except Exception as e:
            print(f"[GSTR-9C reader] Error while computing Table11: {e}")

        # Part 26. Table 12F - Unreconciled ITC - Return value if + ve else zero
        try:
            table_12 = useful_tables[table_position_in_useful_tables_gstr9c["Table_12"]]
            # table_12_F = table_12.iloc[len(table_12) - 1, 2]  # Row F is last row
            table_12_F = table_12[table_12.iloc[:, 0] == 'F']
            valuesFrom9c["table_12_F"] = table_12_F.iloc[0, 2]
            print(f"[GSTR-9C] Table 12F - unreconciled ITC: {valuesFrom9c['table_12_F']}")
        except Exception as e:
            print(f"[GSTR-9C reader] Error while computing Table12 F: {e}")

        # Part 27. Table 16- Amount Payable due to ITC reconciliation -
        # Amount payable - CGST, SGST, IGST, Cess, interest, Penalty.
        try:
            table_16 = useful_tables[table_position_in_useful_tables_gstr9c["Table_16"]]
            table_16_sum_total = pd.to_numeric(table_16.iloc[:, 2], errors='coerce').sum(skipna=True)
            print(f"[GSTR-9C] Table 16 - Amount Payable due to ITC reconciliation: {table_16_sum_total}")
            valuesFrom9c["table_16_sum_total"] = table_16_sum_total
            print(" === ✅ Returning after successful execution of file gstr9c_pdf_reader.py ===")
        except Exception as e:
            print(f"[GSTR-9C reader] Error while computing Table16 F: {e}")

        return valuesFrom9c
    except Exception as e:
        print(f"[GSTR-9C reader] ❌ Error: {e}")
        return valuesFrom9c
        # print(" Error raised to parent class.")
        # raise Exception(e)
