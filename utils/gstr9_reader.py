import os
import pandas as pd
from glob import glob
import pdfplumber
from tabulate import tabulate
from utils.globals.constants import int_eighteen

# Define how many rows to skip per table before setting headers. Don't modify this structure else uploaded
# table_position_in_useful_tables will have to be changed accordingly. Out of all tables extracted from the
# PDF file, only the tables whose positions are mentioned as keys in table_header_rows_skip will be saved in
# useful_tables. The table name : position as per GSTR-9 PDF file is stored in table_position_in_useful_tables.
table_header_rows_skip_gstr9_old_format = {
    0: -1,
    1: 4,  # Table 4
    2: 0,  # Table 4
    3: 4,  # Table 5
    4: 0,  # Table 5
    # 5 : 4, not required
    6: 3,  # Table 6
    7: 3,  # Table 7
    # 8: 0,  Part of Table 7 not required
    9: 3,  # Table 8
    10: 3  # Table 9
}
table_header_rows_skip_gstr9_new_format = {
    0: -1,
    1: 4,  # Table 4
    2: 0,  # Table 4
    3: 4,  # Table 5
    4: -1,  # Table 5
    # 5 : 4, not required
    # 6: 0,  not required
    7: 3,  # Table 6
    8: 3,  # Table 7
    # 9 : 0,  Part of Table 7 not required
    10: 3,  # Table 8
    11: 3,  # Table 9
}
table_position_in_useful_tables_gstr9 = {
    "Table_1": 0,
    "Table_4_part_I": 1,
    "Table_4_part_II": 2,
    "Table_5_part_I": 3,
    "Table_5_part_II": 4,
    "Table_6": 5,
    "Table_7_part_I": 6,
    "Table_8": 7,
    "Table_9": 8
}


async def gstr9_reader(gstin):
    print(" === Starting execution of file gstr9_reader.py ===")
    all_tables = []
    useful_tables = []
    valuesFrom9 = {}
    input_path_of_GSTR_9 = f"uploaded_files/{gstin}/GSTR-9/"

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

        if len(all_tables) == int_eighteen:  # Old format = 18, New format = 19 tables
            table_header_rows_skip = table_header_rows_skip_gstr9_old_format
        else:
            table_header_rows_skip = table_header_rows_skip_gstr9_new_format
        # Clean and process specific tables of GSTR-9
        for idx, skip_rows in table_header_rows_skip.items():
            if idx < len(all_tables):
                df = pd.DataFrame(all_tables[idx])
                df.columns = df.iloc[skip_rows]  # Set new header
                df = df.drop(index=range(skip_rows + 1))  # Drop the header rows
                df = df.reset_index(drop=True)  # Reset index
                useful_tables.append(df)
                # Table 6 can't be printed using tabulate due to its merged cell structure
                # print(f"Table no: {idx}")
                # print(tabulate(df, tablefmt='grid', maxcolwidths=20))

        print(f"useful_tables size: {len(useful_tables)}")  # It should be 8 for both old & new.

        # Write the dataframes in excel sheet GSTR-9.xlsx
        output_path_GSTR_9 = f"reports/{gstin}/GSTR-9/GSTR-9.xlsx"
        with pd.ExcelWriter(output_path_GSTR_9, engine="xlsxwriter") as writer:
            for i, df in enumerate(useful_tables):
                sheet_name = f"Table_{i + 1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]
                # Create a text wrap format
                workbook = writer.book
                wrap_format = workbook.add_format({'text_wrap': True})
                # Apply wrap format + set column width 20 for all columns
                worksheet.set_column(0, len(df.columns) - 1, 20, wrap_format)
            print("GSTR-9.xlsx file written successfully...")

        # 1. Table_4_Part_I_row_G
        table_4_part_I = useful_tables[table_position_in_useful_tables_gstr9["Table_4_part_I"]]
        table4_G1 = table_4_part_I.iloc[6, 1] + " charge"
        table4_G2 = table_4_part_I.iloc[6, 2]
        valuesFrom9["table4_G1"] = table4_G1
        valuesFrom9["table4_G2"] = table4_G2
        print("Table_4_Part_I_row_G calculation done.")

        # 2. Table_4_Part_II_row_N
        table_4_part_II = useful_tables[table_position_in_useful_tables_gstr9["Table_4_part_II"]]
        table4_N1 = table_4_part_II.iloc[len(table_4_part_II) - 1, 1]  # Row N is last row in both
        table4_N2 = table_4_part_II.iloc[len(table_4_part_II) - 1, 2]  # the formats new and old.
        valuesFrom9["table4_N1"] = table4_N1
        valuesFrom9["table4_N2"] = table4_N2
        print("Table_4_Part_I_row_N calculation done.")

        # 3. Table_5_Part_I_row_D
        table_5_part_I = useful_tables[table_position_in_useful_tables_gstr9["Table_5_part_I"]]
        table_5_row_D = table_5_part_I[table_5_part_I.iloc[:, 0] == 'D']  # Find the row in table
        table_5_D1 = table_5_row_D.iloc[0, 1]  # where first column is D.
        table_5_D2 = table_5_row_D.iloc[0, 2]
        valuesFrom9["table_5_D1"] = table_5_D1
        valuesFrom9["table_5_D2"] = table_5_D2
        print("Table_5_Part_I_row_D calculation done.")

        # 4. Table_5_row_N
        table_5_part_II = useful_tables[table_position_in_useful_tables_gstr9["Table_5_part_II"]]
        table_5_N1 = table_5_part_II.iloc[len(table_5_part_II) - 1, 1]  # Row N is last row in both
        table_5_N2 = table_5_part_II.iloc[len(table_5_part_II) - 1, 2]  # the formats new and old.
        valuesFrom9["table_5_N1"] = table_5_N1
        valuesFrom9["table_5_N2"] = table_5_N2
        print("Table_5_row_N calculation done.")

        # 5. Table_6_row_H sum
        table_6 = useful_tables[table_position_in_useful_tables_gstr9["Table_6"]]
        row_h = table_6[table_6.iloc[:, 0] == 'H']  # Find row in table where first column is H.
        sum_table6_row_H = pd.to_numeric(row_h.iloc[0, 2:], errors='coerce').sum(skipna=True)
        valuesFrom9["sum_table6_row_H"] = sum_table6_row_H
        print("Table_6_row_H calculation done.")

        # 6. Table_7_row_A
        table_7 = useful_tables[table_position_in_useful_tables_gstr9["Table_7_part_I"]]
        sum_table7_row_A = pd.to_numeric(table_7.iloc[0, 2:], errors='coerce').sum(skipna=True)
        valuesFrom9["sum_table7_row_A"] = sum_table7_row_A
        print("Table_7_row_A calculation done.")

        # 7. Table_7_row_C
        row_values_7C = table_7.iloc[2, 2:].apply(lambda x: str(x).replace(',', '').strip())
        sum_table7_row_C = pd.to_numeric(row_values_7C, errors='coerce').sum(skipna=True)
        valuesFrom9["sum_table7_row_C"] = sum_table7_row_C
        print("Table_7_row_C calculation done")

        # 8. Table_8_row_D
        table_8 = useful_tables[table_position_in_useful_tables_gstr9["Table_8"]]
        table_8_D1 = table_8.iloc[3, 1]
        table_8_D2 = table_8.iloc[3, 2]
        table_8_D3 = table_8.iloc[3, 3]
        table_8_D4 = table_8.iloc[3, 4]
        table_8_D5 = table_8.iloc[3, 5]
        row_values_8D = table_8.iloc[3, 2:].apply(lambda x: str(x).replace(',', '').strip())
        sum_table8_row_D = pd.to_numeric(row_values_8D, errors='coerce').sum(skipna=True)
        valuesFrom9.update({"table_8_D1": table_8_D1, "table_8_D1": table_8_D1, "table_8_D2": table_8_D2,
                            "table_8_D3": table_8_D3, "table_8_D4": table_8_D4, "table_8_D5": table_8_D5,
                            "sum_table8_row_D": sum_table8_row_D})
        print("Table_8_row_D calculation done")

        # 9. Table_8_row_I
        table_8_I1 = table_8.iloc[8, 1]
        table_8_I2 = table_8.iloc[8, 2]
        table_8_I3 = table_8.iloc[8, 3]
        table_8_I4 = table_8.iloc[8, 4]
        table_8_I5 = table_8.iloc[8, 5]
        row_values_8I = table_8.iloc[8, 2:].apply(lambda x: str(x).replace(',', '').strip())
        sum_table8_row_I = pd.to_numeric(row_values_8I, errors='coerce').sum(skipna=True)
        valuesFrom9.update({"table_8_I1": table_8_I1, "table_8_I2": table_8_I2, "table_8_I3": table_8_I3,
                            "table_8_I4": table_8_I4, "table_8_I5": table_8_I5,
                            "sum_table8_row_I": sum_table8_row_I})
        print("Table_8_row_I calculation done")

        # 10. Table 9: Tax Payable == Paid through cash + Paid through ITC
        table_9 = useful_tables[table_position_in_useful_tables_gstr9["Table_9"]]
        tax_payable_T9 = pd.to_numeric(table_9.iloc[:, 2].str.replace(',', '').str.strip(),
                                       errors='coerce').sum(skipna=True)
        paid_through_cash_T9 = pd.to_numeric(table_9.iloc[:, 3].str.replace(',', '').str.strip(),
                                             errors='coerce').sum(skipna=True)
        paid_through_ITC_CGST_T9 = pd.to_numeric(table_9.iloc[:, 4].str.replace(',', '').str.strip(),
                                                 errors='coerce').sum(skipna=True)
        paid_through_ITC_SGST_T9 = pd.to_numeric(table_9.iloc[:, 5].str.replace(',', '').str.strip(),
                                                 errors='coerce').sum(skipna=True)
        paid_through_ITC_IGST_T9 = pd.to_numeric(table_9.iloc[:, 6].str.replace(',', '').str.strip(),
                                                 errors='coerce').sum(skipna=True)
        paid_through_ITC_Cess_T9 = pd.to_numeric(table_9.iloc[:, 7].str.replace(',', '').str.strip(),
                                                 errors='coerce').sum(skipna=True)
        paid_through_ITC_T9 = paid_through_ITC_CGST_T9 + paid_through_ITC_SGST_T9 + paid_through_ITC_IGST_T9 + paid_through_ITC_Cess_T9
        valuesFrom9.update({"tax_payable_T9": tax_payable_T9, "paid_through_cash_T9": paid_through_cash_T9,
                            "paid_through_ITC_CGST_T9": paid_through_ITC_CGST_T9,
                            "paid_through_ITC_SGST_T9": paid_through_ITC_IGST_T9, "table_8_I5": table_8_I5,
                            "paid_through_ITC_IGST_T9": paid_through_ITC_IGST_T9,
                            "paid_through_ITC_Cess_T9": paid_through_ITC_Cess_T9,
                            "paid_through_ITC_T9": paid_through_ITC_T9})
        print("Table 9 calculation done")
        print(" === ✅ Returning after successful execution 0f file gstr9_reader.py ===")
        return valuesFrom9
    except Exception as e:
        print(f"[GSTR-9 reader] ❌ Error: {e}")
        print(" Error raised to parent class.")
        raise Exception(e)
