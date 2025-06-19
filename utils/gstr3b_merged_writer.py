import os
import pandas as pd
from glob import glob
from collections import defaultdict
from utils.extractors.gstr3b_table_extractor import extract_fixed_tables_from_gstr3b
from utils.globals.constants import newFormat, str_six_point_one, str_two, str_one, \
    str_three_point_one_point_one, oldFormat, parse_month_year, clean_and_parse_number, result_point_9
import datetime
from dateutil.relativedelta import relativedelta

manual_columns = [
    "Description",
    "Total Tax Payable",
    "Integrated Tax paid through ITC",
    "Central Tax paid through ITC",
    "State/UT Tax paid through ITC",
    "Cess paid through ITC",
    "Tax paid in cash",
    "Interest paid in cash",
    "Late fee paid in cash"
]
interest_calc_headers = [
    "Financial Year", "Return Month", "Date of ARN", "Due Date", "Days Late", "Tax paid in cash",
    "Calculated Interest", "Interest paid in cash", "Interest Due"
]


async def generate_gstr3b_merged(input_dir, output_dir):
    final_result_points = {}
    print("=== Generating GSTR-3B merged report ===")
    try:
        gstr3b_format = newFormat  # Let by-default be NEW_FORMAT

        pdf_files = glob(os.path.join(input_dir, "*.pdf"))
        if not pdf_files:
            raise FileNotFoundError("No PDF files found in input directory.")

        print(f"Found {len(pdf_files)} PDF files.")
        interest_matrix = []  # A list of lists: each element list contains table 1, 2, 6.1 for interest calculation
        # For a given key ("3.1"), combined_tables contains all the 3.1 tables from multiple uploaded PDF files
        combined_tables = defaultdict(list)
        for pdf_path in pdf_files:
            tables_list = []
            table_map = extract_fixed_tables_from_gstr3b(pdf_path)
            for key, df in table_map.items():
                combined_tables[key].append(df)  # contains tables from all PDF files with table number as key
                if key in (str_one, str_two, str_six_point_one):  # We need tables 1,2 & 6 for interest calc.
                    tables_list.append(df)
            interest_matrix.append(tables_list)  # Used in Interest Calculation sheet
        # Table 3.1.1 is available only in new format
        if str_three_point_one_point_one not in combined_tables:
            gstr3b_format = oldFormat

        final_tables = {}
        for key, df_list in combined_tables.items():
            if key in (str_one, str_two):  # We are not adding values from two tables as these are info values
                final_tables[key] = df_list[0]
                continue
            elif key == str_six_point_one:
                preprocess_table_6(df_list)
            base_df = df_list[0].copy(deep=True)
            # print(f"\nProcessing table: {key}")
            # print(f"Number of files: {len(df_list)}")
            for row_idx in range(base_df.shape[0]):  # Summation logic cell by cell
                for col_idx in range(1, base_df.shape[1]):
                    total = 0.0
                    for df_num, df in enumerate(df_list):
                        try:
                            num = clean_and_parse_number(df.iat[row_idx, col_idx])
                            #     print(f"File {df_num}: Cell[{row_idx},{col_idx}] = {num}")
                            total += num
                            #     print(f"  File {df_num}: Cell[{row_idx},{col_idx}] = {val} → {num}")
                        except Exception as e:
                            print(f"  [Error] File {df_num}: Cell[{row_idx},{col_idx}] → {e}")
                            continue
                    base_df.iat[row_idx, col_idx] = total  # if pd.notnull(total) and total != 0 else ""
            final_tables[key] = base_df  # Contains summed up values of tables to be written in excel

        # Start with interest calculation
        interest_records = calculate_interest(interest_matrix)
        # Some of only those values where interest due is +ve
        positive_interest_due = sum(row[-1] for row in interest_records if row[-1] > 0)
        final_result_points[result_point_9] = positive_interest_due
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "GSTR-3B_merged.xlsx")
        print(f"Starting to write GSTR-3B_merged sheet to: {output_path}")
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            start_row = 1  # Reserve row 0 for format info ("Old Format" or "New Format")
            sheet_name = "GSTR-3B_merged"

            for key, df in final_tables.items():
                title_df = pd.DataFrame([[f"Table {key}"]])
                title_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False, header=False)
                start_row += 1
                df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
                start_row += len(df) + 2

            # Access the worksheet and workbook
            worksheet = writer.sheets[sheet_name]
            workbook = writer.book  # ✅ Correct place to call add_format
            # Create a wrapped cell format
            wrap_format = workbook.add_format({'text_wrap': True})
            # Write format info in cell A1
            worksheet.write(0, 0, gstr3b_format, wrap_format)  # gstr3b_format = "Old Format" or "New Format"
            # Apply wrap format and width to all relevant columns
            num_columns = max(len(df.columns) for df in final_tables.values())
            worksheet.set_column(0, num_columns - 1, 30, wrap_format)
            print(f"Completed writing GSTR-3B_merged sheet to: {output_path}")

            # ✅ NEW: Write Interest Calculation sheet
            print(f"Starting to write Interest Calculation sheet to: {output_path}")
            interest_df = pd.DataFrame(interest_records, columns=interest_calc_headers)
            interest_df.to_excel(writer, sheet_name="Interest Calculation", index=False)
            interest_sheet = writer.sheets["Interest Calculation"]
            interest_sheet.set_column(0, len(interest_df.columns) - 1, 30, wrap_format)
            print(f"Completed writing Interest Calculation sheet to: {output_path}")

        print(f"✅ GSTR-3B_merged.xlsx saved to: {output_path}")
        return output_path, final_result_points  # ✅ Return the file path for use in API response
    except Exception as e:
        print(f"[GSTR3b_merged]: ❌ Error while merging GSTR-3B files. {e}")
        return output_path, final_result_points

def calculate_interest(interest_matrix):
    print(f"Starting with interest calculation, length of Interest matrix: {len(interest_matrix)}")
    records = []
    for entry in interest_matrix:
        try:
            first_table, second_table, table_6_1 = entry
            table_6_1 = preprocess_table_6([table_6_1])[0]
            # Extract financial year and month
            financial_year = first_table.iloc[0, 1]
            return_month = first_table.iloc[1, 1]
            print(f"Calculating interest for: {return_month, financial_year}")
            # Extract filing date
            filing_date_str = second_table.iloc[-1, -1]
            filing_date = datetime.datetime.strptime(filing_date_str, "%d/%m/%Y").date()
            # Get due date = 20th of next month from return month
            return_month_date = parse_month_year(return_month, financial_year)  # Error if month like July-Sep
            due_date = (return_month_date + relativedelta(months=1)).replace(day=20)
            days_late = max((filing_date - due_date).days, 0)
            # Sum 7th, 8th and 9th columns (Tax paid in cash, Interest paid in cash, Late fee paid in cash)
            tax_paid_in_cash = pd.to_numeric(table_6_1.iloc[:, 6], errors='coerce').fillna(0).sum()
            interest_paid_in_cash = pd.to_numeric(table_6_1.iloc[:, 7], errors='coerce').fillna(0).sum()
            # late_fee_paid_in_cash = pd.to_numeric(table_6_1.iloc[:, 8], errors='coerce').fillna(0).sum()
            calculated_interest = ((tax_paid_in_cash * days_late) / 365) * 0.18 if days_late > 0 else 0
            interest_due = calculated_interest - interest_paid_in_cash

            records.append([
                financial_year,
                return_month,
                filing_date.strftime("%d-%m-%Y"),
                due_date.strftime("%d-%m-%Y"),
                days_late,
                round(tax_paid_in_cash, 2),
                round(calculated_interest, 2),
                round(interest_paid_in_cash, 2),
                # round(late_fee_paid_in_cash, 2),
                round(interest_due, 2)
            ])
        except Exception as e:
            print(f"[GSTR3b_merged]: ❌ Error while calculating interest for: {return_month, financial_year}")
            print(e)
            records.append([
                financial_year,
                return_month,
                filing_date.strftime("%d-%m-%Y")])

    return records


def preprocess_table_6(df_list):
    """ Cleans and standardizes the structure of GSTR-3B Table 6.1 across multiple files.
        Assumes df_list is a list of DataFrames for Table 6.1.
        """
    for i in range(len(df_list)):
        df_list[i].columns = manual_columns
        df_list[i] = df_list[i].iloc[1:].reset_index(drop=True)
    return df_list
