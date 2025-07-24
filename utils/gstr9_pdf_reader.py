import os
import pandas as pd
from glob import glob
import pdfplumber
import datetime
from tabulate import tabulate
from utils.globals.constants import int_eighteen, clean_and_parse_number, newFormat, oldFormat

financial_year_2019_20 = "2019-20"
financial_year_2020_21 = "2020-21"

# Define how many rows to skip per table before setting headers. Don't comment out any line else uploaded
# table_position_in_useful_tables will have to be changed accordingly. Out of all tables extracted from the
# PDF file, only the tables whose positions are mentioned as keys in table_header_rows_skip will be saved in
# useful_tables. The table name : position as per GSTR-9 PDF file is stored in table_position_in_useful_tables.
table_header_rows_skip_gstr9_old_format = {
    0: -1,  # We are not dropping any row
    1: 4,  # Table 4 Part I : here 4 means We are skipping first 5 header rows
    2: -1,  # Table 4 Part II  : 0 means We are skipping first row only
    3: 4,  # Table 5 Part I
    4: -1,  # Table 5 Part II
    # 5 : 4, Table 6 part I not required
    6: -1,  # Table 6 part II
    7: 3,  # Table 7
    # 8: 0,  Part II of Table 7 not required
    9: 3,  # Table 8
    10: 3,  # Table 9
    11: 2  # Table 10, 11, 12 ,13
}
table_header_rows_skip_gstr9_new_format = {
    0: -1,
    1: 4,  # Table 4
    2: -1,  # Table 4
    3: 4,  # Table 5 Part I
    4: -1,  # Table 5 Part II
    # 5 : 3, Table 6 Part I not required
    # 6: 0,  Table 6 Part II not required
    7: -1,  # Table 6 Part III
    8: 3,  # Table 7
    # 9 : 0,  Part of Table 7 not required
    10: 3,  # Table 8
    11: 3,  # Table 9
    12: 2  # Table 10, 11, 12 ,13
}
header_row_map_old_format = {
    # 0:
    1: 2,
    3: 2,
    7: 1,
    9: 1,
    10: 2,
    11: 1
}
header_row_map_new_format = {
    # 0:
    1: 2,
    3: 2,
    8: 1,
    10: 1,
    11: 2,
    12: 1
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
    "Table_9": 8,
    "Table_10_11_12_13": 9
}


async def gstr9_pdf_reader(gstin):
    print(f"[GSTR-9 reader] Starting execution of file gstr9_pdf_reader.py ===")
    all_tables = []
    useful_tables = []
    valuesFrom9 = {}
    gstr9_format = oldFormat  # Let by-default be OLD_FORMAT
    input_path_of_GSTR_9 = f"uploaded_files/{gstin}/GSTR-9/"
    output_path_GSTR_9 = f"reports/{gstin}/GSTR-9.xlsx"

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
            header_row_map = header_row_map_old_format
            print(f"GSTR-9.pdf is based on old format")
        else:
            gstr9_format = newFormat
            table_header_rows_skip = table_header_rows_skip_gstr9_new_format
            header_row_map = header_row_map_new_format
            print(f"GSTR-9.pdf is based on new format")

        # Clean and process specific tables of GSTR-9
        for idx, skip_rows in table_header_rows_skip.items():
            if idx < len(all_tables):
                df = pd.DataFrame(all_tables[idx])
                if idx in header_row_map:
                    df.columns = df.iloc[header_row_map.get(idx)]  # Set new header
                df = df.drop(index=range(skip_rows + 1))  # Drop the header rows
                df = df.reset_index(drop=True)  # Reset index
                useful_tables.append(df)
                # Table 6 can't be printed using tabulate due to its merged cell structure
                #     print(f"Table no: {idx}")
                #     print(tabulate(df, tablefmt='grid', maxcolwidths=20))

        # Set column headers which are not set
        setColumnHeaders(useful_tables)

        # Clean the values of useful_tables content before saving to excel.
        for df in useful_tables:
            for row_idx in range(df.shape[0]):
                for col_idx in range(2, df.shape[1]):
                    df.iat[row_idx, col_idx] = clean_and_parse_number(df.iat[row_idx, col_idx])

        print(f"useful_tables size: {len(useful_tables)}")  # It should be 10 for both old & new.
        table_4_merged = pd.concat([useful_tables[table_position_in_useful_tables_gstr9["Table_4_part_I"]],
                                    useful_tables[table_position_in_useful_tables_gstr9["Table_4_part_II"]]],
                                   ignore_index=True)
        table_5_merged = pd.concat([useful_tables[table_position_in_useful_tables_gstr9["Table_5_part_I"]],
                                    useful_tables[table_position_in_useful_tables_gstr9["Table_5_part_II"]]],
                                   ignore_index=True)
        # useful_tables[table_position_in_useful_tables_gstr9["Table_4_part_I"]] = table_4_merged
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
            print(f"[GSTR-9 reader] GSTR-9.xlsx file written successfully...")

        # 1. Table_4_Part_1_row_G (CGST + SGST + IGST + Cess)
        try:
            table4_G1 = "Total tax on inward supplies liable for reverse charge"
            # table4_G2 = pd.to_numeric(table_4_part_I.iloc[6, 3:], errors='coerce').sum(skipna=True)
            valuesFrom9["table4_G1"] = table4_G1
            row_G = table_4_merged[table_4_merged.iloc[:, 0] == 'G']
            valuesFrom9["table4_G2"] = pd.to_numeric(row_G.iloc[0, 3:], errors='coerce').sum(skipna=True)
            print(f"Table_4_Part_I_row_G calculation done: {valuesFrom9['table4_G2']}")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table4 G: {e}")

        # 2. Table_4_Part_II_row_N (CGST + SGST + IGST + Cess)
        try:
            # table_4_part_II = useful_tables[table_position_in_useful_tables_gstr9["Table_4_part_II"]]
            # table4_N2 = pd.to_numeric(table_4_merged.iloc[len(table_4_merged) - 1, 3:],  # Row N is last row
            #                           errors='coerce').sum(skipna=True)  # in both the formats new and old.
            table4_N1 = "Total tax on Supplies and advances"
            valuesFrom9["table4_N1"] = table4_N1
            row_N = table_4_merged[table_4_merged.iloc[:, 0] == 'N']
            table4_N2 = pd.to_numeric(row_N.iloc[0, 3:], errors='coerce').sum(skipna=True)
            valuesFrom9["table4_N2"] = table4_N2
            print(f"Table_4_Part_II_row_N calculation done: {table4_N2}")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table4 N: {e}")

        # 3. Table_5_Part_1 0r Part_2 row_D_&_E
        try:
            # table_5_part_I = useful_tables[table_position_in_useful_tables_gstr9["Table_5_part_I"]]
            # table_5_part_II = useful_tables[table_position_in_useful_tables_gstr9["Table_5_part_II"]]
            table_5_row_D = table_5_merged[table_5_merged.iloc[:, 0] == 'D']  # Find the row in table where first column is D.
            table_5_row_E = table_5_merged[table_5_merged.iloc[:, 0] == 'E']  # Find the row in table where first column is E.
            table_5_D1 = "Taxable value (Exempted + Nil rated)"
            table_5_D2 = table_5_row_D.iloc[0, 2] + table_5_row_E.iloc[0, 2]
            valuesFrom9["table_5_D1"] = table_5_D1
            valuesFrom9["table_5_D2"] = table_5_D2
            print(f"Table_5_Part_I_row_D_&_E calculation done: {table_5_D2}")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table5 D: {e}")

        # 4. Table_5_Part_II_row_N
        try:
            # table_5_part_II = useful_tables[table_position_in_useful_tables_gstr9["Table_5_part_II"]]
            table_5_row_N = table_5_merged[table_5_merged.iloc[:, 0] == 'N']
            table_5_N1 = table_5_row_N.iloc[0, 1]  # Row N is last row in both
            table_5_N2 = table_5_row_N.iloc[0, 2]  # the formats new and old.
            valuesFrom9["table_5_N1"] = table_5_N1
            valuesFrom9["table_5_N2"] = table_5_N2
            print(f"Table_5_Part_II_row_N calculation done: {table_5_N2}")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table5 N: {e}")

        #  Parameter 13 of ASMT-10 report: calculate late fee for GSTR-9 late filing
        # For FY 2019-20 = 31st Mar 2021
        # For FY 2020-21 = 28th Feb 2022
        # Else 31st Dec of that year
        try:
            table1 = useful_tables[table_position_in_useful_tables_gstr9["Table_1"]]
            financial_year = table1.iloc[0, 1]
            filing_date_str = table1.iloc[-1, -1]
            if financial_year == financial_year_2019_20:
                print(f"[GSTR-9]Late fee calculation: Setting due day as 31st Mar 2021 for year: {financial_year}")
                due_date_for_ITC = datetime.datetime.strptime("31/03/2021", "%d/%m/%Y").date()
            elif financial_year == financial_year_2020_21:
                print(f"[GSTR-9]Late fee calculation: Setting due day as 28th Feb 2022 for year: {financial_year}")
                due_date_for_ITC = datetime.datetime.strptime("28/02/2022", "%d/%m/%Y").date()
            else:
                fy_end_year = int(financial_year.split('-')[0]) + 1  # 2022-23 becomes 2023
                due_day = f"31/12/{fy_end_year}"
                print(f"[GSTR-9]Late fee calculation: Setting due day as {due_day} for year: {financial_year}")
                due_date_for_ITC = datetime.datetime.strptime(due_day, "%d/%m/%Y").date()
            print(f"Due date: {due_date_for_ITC}")
            filing_date = datetime.datetime.strptime(filing_date_str, "%d-%m-%Y").date()
            if filing_date > due_date_for_ITC and table_5_N2 >= 20000000:
                days_late = max((filing_date - due_date_for_ITC).days, 0)
                late_fee_by_days = (100 * days_late)
                late_fee_by_quarter_percent_of_5N = round(0.0025 * table_5_N2, 2)
                print(f"Late_fee_by_days: {late_fee_by_days}")
                print(f"Late_fee_by_quarter_percent_of_5N: {late_fee_by_quarter_percent_of_5N}")
                late_fee_gstr9_applicable = max(late_fee_by_days, late_fee_by_quarter_percent_of_5N)
            else:
                late_fee_gstr9_applicable = 0.00
            valuesFrom9['late_fee_gstr9_applicable'] = late_fee_gstr9_applicable
            print(f" Late fee GSTR-9 calculation done: {late_fee_gstr9_applicable}")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing late fee (point 13) of ASMT-10 report: {e}")

        # 5. Table_6_row_H sum
        try:
            table_6 = useful_tables[table_position_in_useful_tables_gstr9["Table_6"]]
            row_h = table_6[table_6.iloc[:, 0] == 'H']  # Find row in table where first column is H.
            valuesFrom9["table6_row_H_CGST"] = pd.to_numeric(row_h.iloc[0, 2], errors='coerce')
            valuesFrom9["table6_row_H_SGST"] = pd.to_numeric(row_h.iloc[0, 3], errors='coerce')
            valuesFrom9["table6_row_H_IGST"] = pd.to_numeric(row_h.iloc[0, 4], errors='coerce')
            valuesFrom9["table6_row_H_CESS"] = pd.to_numeric(row_h.iloc[0, 5], errors='coerce')
            print(f"Table_6_row_H calculation done.")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table_6_row_H: {e}")

        # 6. Table_7_row_A
        try:
            table_7 = useful_tables[table_position_in_useful_tables_gstr9["Table_7_part_I"]]
            valuesFrom9["table7_row_A_CGST"] = pd.to_numeric(table_7.iloc[0, 2], errors='coerce')
            valuesFrom9["table7_row_A_SGST"] = pd.to_numeric(table_7.iloc[0, 3], errors='coerce')
            valuesFrom9["table7_row_A_IGST"] = pd.to_numeric(table_7.iloc[0, 4], errors='coerce')
            valuesFrom9["table7_row_A_CESS"] = pd.to_numeric(table_7.iloc[0, 5], errors='coerce')
            print(f"Table_7_row_A Rule 37 calculation done.")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table_7_row_A Rule 37: {e}")

        # 7. Table_7_row_C
        # row_values_7C = table_7.iloc[2, 2:].apply(lambda x: str(x).replace(',', '').strip())
        # sum_table7_row_C = pd.to_numeric(row_values_7C, errors='coerce').sum(skipna=True)
        try:
            sum_table7_row_C = pd.to_numeric(table_7.iloc[2, 2:], errors='coerce').sum(skipna=True)
            valuesFrom9["sum_table7_row_C"] = sum_table7_row_C
            print(f"Table_7_row_C Rule 42 calculation done: {sum_table7_row_C} ")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table_7_row_C: {e}")

        # 8. Table_8_row_D
        try:
            table_8 = useful_tables[table_position_in_useful_tables_gstr9["Table_8"]]
            table_8_D1 = table_8.iloc[3, 1]
            table_8_D2 = table_8.iloc[3, 2]
            table_8_D3 = table_8.iloc[3, 3]
            table_8_D4 = table_8.iloc[3, 4]
            table_8_D5 = table_8.iloc[3, 5]
            sum_table8_row_D = pd.to_numeric(table_8.iloc[3, 2:], errors='coerce').sum(skipna=True)
            valuesFrom9.update({"table_8_D1": table_8_D1, "table_8_D2": table_8_D2,
                                "table_8_D3": table_8_D3, "table_8_D4": table_8_D4, "table_8_D5": table_8_D5,
                                "sum_table8_row_D": sum_table8_row_D})
            print(f"Table_8_row_D calculation done: {sum_table8_row_D}")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table_8_row_D: {e}")

        # 9. Table_8_row_I
        try:
            table_8_I1 = table_8.iloc[8, 1]
            table_8_I2 = table_8.iloc[8, 2]
            table_8_I3 = table_8.iloc[8, 3]
            table_8_I4 = table_8.iloc[8, 4]
            table_8_I5 = table_8.iloc[8, 5]
            sum_table8_row_I = pd.to_numeric(table_8.iloc[8, 2:], errors='coerce').sum(skipna=True)
            valuesFrom9.update({"table_8_I1": table_8_I1, "table_8_I2": table_8_I2, "table_8_I3": table_8_I3,
                                "table_8_I4": table_8_I4, "table_8_I5": table_8_I5,
                                "sum_table8_row_I": sum_table8_row_I})
            print(f"Table_8_row_I calculation done: {sum_table8_row_I}")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table_8_row_I: {e}")

        # 10. Table 9: Tax Payable == Paid through cash + Paid through ITC
        try:
            table_9 = useful_tables[table_position_in_useful_tables_gstr9["Table_9"]]
            tax_payable_T9 = pd.to_numeric(table_9.iloc[:, 2], errors='coerce').sum(skipna=True)
            paid_through_cash_T9 = pd.to_numeric(table_9.iloc[:, 3], errors='coerce').sum(skipna=True)
            paid_through_ITC_CGST_T9 = pd.to_numeric(table_9.iloc[:, 4], errors='coerce').sum(skipna=True)
            paid_through_ITC_SGST_T9 = pd.to_numeric(table_9.iloc[:, 5], errors='coerce').sum(skipna=True)
            paid_through_ITC_IGST_T9 = pd.to_numeric(table_9.iloc[:, 6], errors='coerce').sum(skipna=True)
            paid_through_ITC_Cess_T9 = pd.to_numeric(table_9.iloc[:, 7], errors='coerce').sum(skipna=True)
            paid_through_ITC_T9 = paid_through_ITC_CGST_T9 + paid_through_ITC_SGST_T9 + paid_through_ITC_IGST_T9 + paid_through_ITC_Cess_T9
            valuesFrom9.update({"tax_payable_T9": tax_payable_T9, "paid_through_cash_T9": paid_through_cash_T9,
                                "paid_through_ITC_CGST_T9": paid_through_ITC_CGST_T9,
                                "paid_through_ITC_SGST_T9": paid_through_ITC_SGST_T9, "table_8_I5": table_8_I5,
                                "paid_through_ITC_IGST_T9": paid_through_ITC_IGST_T9,
                                "paid_through_ITC_Cess_T9": paid_through_ITC_Cess_T9,
                                "paid_through_ITC_T9": paid_through_ITC_T9})
            print(f"Table 9 calculation done: 1. Tax payable: {tax_payable_T9} "
                  f"2. Paid through cash: {paid_through_cash_T9} 3. Paid through ITC: {paid_through_ITC_T9}")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table 9: Tax Payable: {e}")

        # 10. Tax payable table 9.
        try:
            valuesFrom9["tax_payable_table9_IGST"] = pd.to_numeric(table_9.iloc[0, 2], errors='coerce')
            valuesFrom9["tax_payable_table9_CGST"] = pd.to_numeric(table_9.iloc[1, 2], errors='coerce')
            valuesFrom9["tax_payable_table9_SGST"] = pd.to_numeric(table_9.iloc[2, 2], errors='coerce')
            valuesFrom9["tax_payable_table9_CESS"] = pd.to_numeric(table_9.iloc[3, 2], errors='coerce')
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Tax payable table 9: {e}")
        # print(f"Tax payable Table 9 IGST: {valuesFrom9['tax_payable_table9_IGST']}")

        # 11. Table 13 (IGST, CGST, SGST, Cess)
        try:
            table_10_11_12_13 = useful_tables[table_position_in_useful_tables_gstr9["Table_10_11_12_13"]]
            valuesFrom9["table_13_1"] = table_10_11_12_13.iloc[3, 1]
            valuesFrom9["table_13_CGST"] = table_10_11_12_13.iloc[3, 3]
            valuesFrom9["table_13_SGST"] = table_10_11_12_13.iloc[3, 4]
            valuesFrom9["table_13_IGST"] = table_10_11_12_13.iloc[3, 5]
            valuesFrom9["table_13_CESS"] = table_10_11_12_13.iloc[3, 6]
            print(f"Table 13: CGST = {table_10_11_12_13.iloc[3, 3]}")
        except Exception as e:
            print(f"[GSTR-9 reader] Error while computing Table 13 (IGST, CGST, SGST, Cess): {e}")

        print(" === ✅ Returning after successful execution of file gstr9_pdf_reader.py ===")
        return valuesFrom9
    except Exception as e:
        print(f"[GSTR-9 reader] ❌ Error: {e}")
        return valuesFrom9


def setColumnHeaders(useful_tables):
    table1 = useful_tables[1]
    table1.columns.values[0] = "Sr.No"
    table1.columns.values[1] = "Nature of Supplies"
    table1.columns.values[2] = "Taxable Value(₹)"
    table2 = useful_tables[2]
    table2.columns = table1.columns
    table3 = useful_tables[3]
    table3.columns.values[0] = "Sr.No"
    table3.columns.values[1] = "Nature of Supplies"
    table3.columns.values[2] = "Taxable Value(₹)"
    table4 = useful_tables[4]
    table4.columns = table3.columns
    table8 = useful_tables[8]
    table8.columns.values[0] = "9"
    table8.columns.values[1] = "Description"
    table8.columns.values[2] = "Tax Payable (₹)"
    table8.columns.values[3] = "Paid Through Cash  (₹)"
