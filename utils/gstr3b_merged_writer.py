import datetime
import os
from collections import defaultdict
from glob import glob

import pandas as pd
from dateutil.relativedelta import relativedelta

from utils.extractors.gstr3b_table_extractor import extract_fixed_tables_from_gstr3b
from utils.globals.constants import newFormat, str_six_point_one, str_two, str_one, \
    str_three_point_one_point_one, oldFormat, parse_month_year, clean_and_parse_number, late_fee_headers, str_four, \
    str_three_point_one, financial_year_2022_23, parse_month, financial_year_2023_24, financial_year_2024_25

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
financial_year_2019_20 = "2019-20"


async def generate_gstr3b_merged(input_dir, output_dir):
    final_result_points = {}
    output_path = None
    print(f"[GSTR-3b_merged_writer] Generating GSTR-3B merged report ===")
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
            interest_tables_list = []
            table_map = extract_fixed_tables_from_gstr3b(pdf_path)
            for key, df in table_map.items():
                combined_tables[key].append(df)  # contains tables from all PDF files with table number as key
                # We need tables 1,2,4 & 6 for interest calc & late fee.
                if key in (str_one, str_two, str_three_point_one, str_four, str_six_point_one):
                    interest_tables_list.append(df)
            interest_matrix.append(interest_tables_list)  # Used in Interest Calculation sheet

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
        print("[GSTR-3B_merged_writer]: Addition of values of tables across all files completed.")

        # Result point 8: Find Ineligible ITC due to delay in filing GSTR-3B
        month_wise_table_4_df, ineligible_ITC_list = calculate_ineligible_ITC(interest_matrix)
        print(f"[GSTR-3B_merged] Total ineligible ITC due to late filing: {ineligible_ITC_list}")
        final_result_points["result_point_8_IGST"] = round(ineligible_ITC_list[0], 2)
        final_result_points["result_point_8_CGST"] = round(ineligible_ITC_list[1], 2)
        final_result_points["result_point_8_SGST"] = round(ineligible_ITC_list[2], 2)
        final_result_points["result_point_8_CESS"] = round(ineligible_ITC_list[3], 2)

        # Result point 9: Start  interest calculation
        interest_records_for_excel, interest_list = calculate_interest(interest_matrix)
        final_result_points['result_point_9_IGST'] = round(interest_list[0], 2)
        final_result_points['result_point_9_CGST'] = round(interest_list[1], 2)
        final_result_points['result_point_9_SGST'] = round(interest_list[2], 2)
        final_result_points['result_point_9_CESS'] = round(interest_list[3], 2)

        # Result point 12: Start late fee calculation
        late_fee_records = calculate_late_fee(interest_matrix)
        total_late_fee = sum(row[-1] for row in late_fee_records)
        print(f"[GSTR-3B] Total late fee: {total_late_fee}")
        final_result_points["result_point_12_total_late_fee_gstr3b"] = total_late_fee
        # Find late fee paid in cash
        table_6_1_merged = final_tables["6.1"]
        late_fee_paid_in_cash = pd.to_numeric(table_6_1_merged.iloc[:, 8], errors='coerce').fillna(0).sum()
        print(f"[GSTR-3B] Total Late fee paid in cash: {late_fee_paid_in_cash}")
        final_result_points["result_point_12_late_fee_paid_in_cash"] = round(late_fee_paid_in_cash, 2)

        # Result point 22: Find cash liability due to less cash payment
        cash_liability = calculate_cash_liability(interest_matrix)
        print(f"Total cash liability: {cash_liability}")
        final_result_points["result_point_22_IGST"] = round(cash_liability[0], 2)
        final_result_points["result_point_22_CGST"] = round(cash_liability[1], 2)
        final_result_points["result_point_22_SGST"] = round(cash_liability[2], 2)
        final_result_points["result_point_22_CESS"] = round(cash_liability[3], 2)

        # Start saving analysis data to excel
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
            interest_df = pd.DataFrame(interest_records_for_excel, columns=interest_calc_headers)
            interest_df.to_excel(writer, sheet_name="Interest Calculation", index=False)
            interest_sheet = writer.sheets["Interest Calculation"]
            interest_sheet.set_column(0, len(interest_df.columns) - 1, 20, wrap_format)
            print(f"Completed writing Interest Calculation sheet to: {output_path}")

            # ✅ Write Late fee sheet
            print(f"Starting to write Late Fee sheet to: {output_path}")
            late_fee_df = pd.DataFrame(late_fee_records, columns=late_fee_headers)
            late_fee_df.to_excel(writer, sheet_name="Late Fee Record", index=False)
            late_fee_sheet = writer.sheets["Late Fee Record"]
            late_fee_sheet.set_column(0, len(late_fee_df.columns) - 1, 20, wrap_format)
            print(f"Completed writing Late Fee sheet to: {output_path}")

            # ✅ Write Month-wise ITC (Table 4) sheet
            print(f"Starting to write month-wise Table_4: {output_path}")
            start_row = 0
            spacing = 2  # Number of blank rows between monthly blocks
            for itc_df in month_wise_table_4_df:
                itc_df.to_excel(writer, sheet_name='Monthly ITC', index=False, startrow=start_row)
                start_row += len(itc_df) + spacing  # Move to next block
            monthly_itc_sheet = writer.sheets["Monthly ITC"]
            monthly_itc_sheet.set_column(0, len(itc_df.columns) - 1, 25, wrap_format)
            print(f"Completed writing Monthly ITC sheet to: {output_path}")

        print(f"✅ GSTR-3B_merged.xlsx saved to: {output_path}")
        return output_path, final_result_points  # ✅ Return the file path for use in API response
    except Exception as e:
        print(f"[GSTR3b_merged]: ❌ Error while merging GSTR-3B files. {e}")
        return output_path, final_result_points


#  Parameter 8 of ASMT-10 report
# Due date for ineligible ITC for financial_year 2019-20 is = 30-Nov-2021
# Due date for ineligible ITC for financial_year 2020-21 is = 30-Nov-2021
# Due date for ineligible ITC for financial_year 2021-22 is = 30-Nov-2022
# Due date for ineligible ITC for financial_year 2022-23 is = 30-Nov-2023
# Due date for ineligible ITC for financial_year 2023-24 is = 30-Nov-2024
# Due date for ineligible ITC for financial_year 2024-25 is = 30-Nov-2025
def calculate_ineligible_ITC(interest_matrix):
    print(f"[GSTR-3B_merged] Starting with ineligible_ITC calculation, number of months : {len(interest_matrix)}")
    monthly_dfs = []
    ineligible_ITC_list = [0.0] * 4  # four elements representing sum of IGST, CGST, SGST, Cess
    # due_date_for_ITC = datetime.datetime.strptime("30/11/2022", "%d/%m/%Y").date()
    for entry in interest_matrix:
        try:
            first_table, second_table, table_3, table_4, table_6_1 = entry
            financial_year = first_table.iloc[0, 1]
            return_month = first_table.iloc[1, 1]
            print(f"Calculating ineligible ITC for: {return_month, financial_year}")
            if financial_year == financial_year_2019_20:
                print(f"[GSTR-3B]Ineligible ITC calculation: Setting due day as 30th Nov 2021 for year: {financial_year}")
                due_date_for_ITC = datetime.datetime.strptime("30/11/2021", "%d/%m/%Y").date()
            else:
                fy_end_year = int(financial_year.split('-')[0]) + 1  # 2022-23 becomes 2023
                due_day = f"30/11/{fy_end_year}"
                print(f"[GSTR-3B]Ineligible ITC calculation: Setting due day as {due_day} for year: {financial_year}")
                due_date_for_ITC = datetime.datetime.strptime(due_day, "%d/%m/%Y").date()
            filing_date_str = second_table.iloc[-1, -1]
            filing_date = datetime.datetime.strptime(filing_date_str, "%d/%m/%Y").date()

            if filing_date > due_date_for_ITC:
                # Sum of rows 1 to 5 of table 4(A) across all columns IGST + CGST + State/UT tax + Cess
                print(f"ITC for month: {return_month, financial_year} is ineligible due to late filing: {filing_date}")
                ineligible_ITC_list[0] += pd.to_numeric(table_4.iloc[1:6, 1], errors='coerce').fillna(
                    0).to_numpy().sum()
                ineligible_ITC_list[1] += pd.to_numeric(table_4.iloc[1:6, 2], errors='coerce').fillna(
                    0).to_numpy().sum()
                ineligible_ITC_list[2] += pd.to_numeric(table_4.iloc[1:6, 3], errors='coerce').fillna(
                    0).to_numpy().sum()
                ineligible_ITC_list[3] += pd.to_numeric(table_4.iloc[1:6, 4], errors='coerce').fillna(
                    0).to_numpy().sum()

            # Add to monthly dataframes
            table_subset = table_4.iloc[1:6, 0:].reset_index(drop=True)
            # Create 'Month' and 'Filing Date' columns (merged-style)
            year_col = [financial_year] + [''] * (table_subset.shape[0] - 1)
            month_col = [return_month] + [''] * (table_subset.shape[0] - 1)
            filing_date_col = [filing_date.strftime("%d-%m-%Y")] + [''] * (table_subset.shape[0] - 1)
            due_date_col = [due_date_for_ITC.strftime("%d-%m-%Y")] + [''] * (table_subset.shape[0] - 1)

            # Combine into a single DataFrame
            monthly_df = pd.DataFrame({
                'Financial Year': year_col,
                'Month': month_col,
                'Filing Date': filing_date_col,
                'Due Date': due_date_col
            }).join(table_subset)
            monthly_dfs.append(monthly_df)
        except Exception as e:
            print(
                f"[GSTR-3b_merged_writer]: ❌ Error while calculating ineligible ITC for: {return_month, financial_year}")
            print(e)
    return monthly_dfs, ineligible_ITC_list


#  Parameter 9 of ASMT-10 report
def calculate_interest(interest_matrix):
    print(f"Starting with interest calculation, length of Interest matrix: {len(interest_matrix)}")
    records = []
    interest_list = [0.0] * 4
    for entry in interest_matrix:
        try:
            first_table, second_table, table_3, table_4, table_6_1 = entry
            table_6_1 = preprocess_table_6([table_6_1])[0]
            # Extract financial year and month
            financial_year = first_table.iloc[0, 1]
            return_month = first_table.iloc[1, 1]
            print(f"Calculating interest for: {return_month, financial_year}")
            # Extract filing date
            filing_date_str = second_table.iloc[-1, -1]
            filing_date = datetime.datetime.strptime(filing_date_str, "%d/%m/%Y").date()
            return_month_date = parse_month_year(return_month, financial_year)
            day = dayOFDue(financial_year, return_month, "Interest")
            due_date = (return_month_date + relativedelta(months=1)).replace(day=day)
            print(f"Due date: {due_date}")
            days_late = max((filing_date - due_date).days, 0)
            # Sum 7th, 8th and 9th columns (Tax paid in cash, Interest paid in cash, Late fee paid in cash)
            tax_paid_in_cash = pd.to_numeric(table_6_1.iloc[:, 6], errors='coerce').fillna(0).sum()
            interest_paid_in_cash = pd.to_numeric(table_6_1.iloc[:, 7], errors='coerce').fillna(0).sum()
            # late_fee_paid_in_cash = pd.to_numeric(table_6_1.iloc[:, 8], errors='coerce').fillna(0).sum()
            calculated_interest = ((tax_paid_in_cash * days_late) / 365) * 0.18 if days_late > 0 else 0
            interest_due = calculated_interest - interest_paid_in_cash

            if days_late > 0:
                igst_tax_paid_in_cash = pd.to_numeric(table_6_1.iloc[[1, 6], 6], errors='coerce').sum(skipna=True)
                cgst_tax_paid_in_cash = pd.to_numeric(table_6_1.iloc[[2, 7], 6], errors='coerce').sum(skipna=True)
                sgst_tax_paid_in_cash = pd.to_numeric(table_6_1.iloc[[3, 8], 6], errors='coerce').sum(skipna=True)
                cess_tax_paid_in_cash = pd.to_numeric(table_6_1.iloc[[4, 9], 6], errors='coerce').sum(skipna=True)

                igst_interest_paid_in_cash = pd.to_numeric(table_6_1.iloc[[1, 6], 7], errors='coerce').sum(skipna=True)
                cgst_interest_paid_in_cash = pd.to_numeric(table_6_1.iloc[[2, 7], 7], errors='coerce').sum(skipna=True)
                sgst_interest_paid_in_cash = pd.to_numeric(table_6_1.iloc[[3, 8], 7], errors='coerce').sum(skipna=True)
                cess_interest_paid_in_cash = pd.to_numeric(table_6_1.iloc[[4, 9], 7], errors='coerce').sum(skipna=True)

                igst_total_interest = ((igst_tax_paid_in_cash * days_late) / 365) * 0.18 if days_late > 0 else 0
                cgst_total_interest = ((cgst_tax_paid_in_cash * days_late) / 365) * 0.18 if days_late > 0 else 0
                sgst_total_interest = ((sgst_tax_paid_in_cash * days_late) / 365) * 0.18 if days_late > 0 else 0
                cess_total_interest = ((cess_tax_paid_in_cash * days_late) / 365) * 0.18 if days_late > 0 else 0

                igst_interest_due = round(igst_total_interest - igst_interest_paid_in_cash, 2)
                cgst_interest_due = round(cgst_total_interest - cgst_interest_paid_in_cash, 2)
                sgst_interest_due = round(sgst_total_interest - sgst_interest_paid_in_cash, 2)
                cess_interest_due = round(cess_total_interest - cess_interest_paid_in_cash, 2)

                interest_list[0] += igst_interest_due if igst_interest_due > 0 else 0
                interest_list[1] += cgst_interest_due if cgst_interest_due > 0 else 0
                interest_list[2] += sgst_interest_due if sgst_interest_due > 0 else 0
                interest_list[3] += cess_interest_due if cess_interest_due > 0 else 0

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
            print(f"[GSTR-3b_merged_writer]: ❌ Error while calculating interest for: {return_month, financial_year}")
            print(e)
            records.append([
                financial_year,
                return_month,
                filing_date.strftime("%d-%m-%Y")])

    return records, interest_list


#  Parameter 12 of ASMT-10 report
def calculate_late_fee(interest_matrix):
    print(f"[GSTR-3B_merged]Starting with late fee calculation, number of months : {len(interest_matrix)}")
    records = []
    for entry in interest_matrix:
        try:
            first_table, second_table, table_3, table_4, table_6_1 = entry
            financial_year = first_table.iloc[0, 1]
            return_month = first_table.iloc[1, 1]
            print(f"Calculating late fee for: {return_month, financial_year}")
            filing_date_str = second_table.iloc[-1, -1]
            filing_date = datetime.datetime.strptime(filing_date_str, "%d/%m/%Y").date()
            day = dayOFDue(financial_year, return_month, "Late Fee")
            return_month_date = parse_month_year(return_month, financial_year)
            due_date = (return_month_date + relativedelta(months=1)).replace(day=day)
            print(f"Due date: {due_date}")
            days_late = max((filing_date - due_date).days, 0)
            calculated_late_fee = (100 * days_late)
            late_fee_applicable = min(calculated_late_fee, 5000)

            records.append([
                financial_year,
                return_month,
                filing_date.strftime("%d-%m-%Y"),
                due_date.strftime("%d-%m-%Y"),
                days_late,
                calculated_late_fee,
                late_fee_applicable
            ])
        except Exception as e:
            print(f"[GSTR-3b_merged_writer]: ❌ Error while calculating late fee for: {return_month, financial_year}")
            print(e)
            records.append([
                financial_year,
                return_month,
                filing_date.strftime("%d-%m-%Y")])
    return records


#  Parameter 22 of ASMT-10 report
def calculate_cash_liability(interest_matrix):
    print(f"Starting with cash liability calculation, length of Interest matrix: {len(interest_matrix)}")
    cash_liability = [0.0] * 4
    for entry in interest_matrix:
        try:
            first_table, second_table, table_3_1, table_4, table_6_1 = entry
            table_6_1 = preprocess_table_6([table_6_1])[0]
            # Extract financial year and month
            financial_year = first_table.iloc[0, 1]
            return_month = first_table.iloc[1, 1]
            print(f"Calculating cash liability for: {return_month, financial_year}")
            taxable_value = pd.to_numeric(table_3_1.iloc[0, 1])
            print(f"table_3 a: taxable value= {taxable_value}")
            if taxable_value > 5000000:
                total_IGST_payable = pd.to_numeric(table_6_1.iloc[1, 1], errors='coerce')
                total_IGST_paid_in_cash = pd.to_numeric(table_6_1.iloc[1, 6], errors='coerce')
                one_percent_of_total_IGST_payable = round(0.01 * total_IGST_payable, 2)
                total_CGST_payable = pd.to_numeric(table_6_1.iloc[2, 1], errors='coerce')
                total_CGST_paid_in_cash = pd.to_numeric(table_6_1.iloc[2, 6], errors='coerce')
                one_percent_of_total_CGST_payable = round(0.01 * total_CGST_payable, 2)
                total_SGST_payable = pd.to_numeric(table_6_1.iloc[3, 1], errors='coerce')
                total_SGST_paid_in_cash = pd.to_numeric(table_6_1.iloc[3, 6], errors='coerce')
                one_percent_of_total_SGST_payable = round(0.01 * total_SGST_payable, 2)
                total_CESS_payable = pd.to_numeric(table_6_1.iloc[4, 1], errors='coerce')
                total_CESS_paid_in_cash = pd.to_numeric(table_6_1.iloc[4, 6], errors='coerce')
                one_percent_of_total_CESS_payable = round(0.01 * total_CESS_payable, 2)
                if total_IGST_paid_in_cash < one_percent_of_total_IGST_payable:
                    print(
                        f"IGST Cash liability for {return_month}, {financial_year} is there due to less cash payment.")
                    cash_liability[0] += round(one_percent_of_total_IGST_payable - total_IGST_paid_in_cash, 2)
                if total_CGST_paid_in_cash < one_percent_of_total_CGST_payable:
                    print(
                        f"CGST Cash liability for {return_month}, {financial_year} is there due to less cash payment.")
                    cash_liability[1] += round(one_percent_of_total_CGST_payable - total_CGST_paid_in_cash, 2)
                if total_SGST_paid_in_cash < one_percent_of_total_SGST_payable:
                    print(
                        f"SGST Cash liability for {return_month}, {financial_year} is there due to less cash payment.")
                    cash_liability[2] += round(one_percent_of_total_SGST_payable - total_SGST_paid_in_cash, 2)
                if total_CESS_paid_in_cash < one_percent_of_total_CESS_payable:
                    print(
                        f"CESS Cash liability for {return_month}, {financial_year} is there due to less cash payment.")
                    cash_liability[3] += round(one_percent_of_total_CESS_payable - total_CESS_paid_in_cash, 2)

        except Exception as e:
            print(f"[GSTR-3b_merged] Error during Cash liability for {return_month}, {financial_year}: {e}")

    return cash_liability


def preprocess_table_6(df_list):
    """ Cleans and standardizes the structure of GSTR-3B Table 6.1 across multiple files.
        Assumes df_list is a list of DataFrames for Table 6.1.
        """
    for i in range(len(df_list)):
        df_list[i].columns = manual_columns
        df_list[i] = df_list[i].iloc[1:].reset_index(drop=True)
    return df_list


# Get due date = 20th of next month from return month
# due date for FY 2022-23, return month April = 24 May 2023
# due date for FY 2023-24, return month May = 30 June 2024
# due date for FY 2024-25, return month Dec = 22 Jan 2025
def dayOFDue(financial_year, return_month, purpose):
    if financial_year == financial_year_2022_23 and parse_month(return_month) == 4:
        print(f"[GSTR-3B] {purpose} calculation: Setting due day as 24th of next month for month: {return_month}, year: {financial_year}")
        day = 24
    elif financial_year == financial_year_2023_24 and parse_month(return_month) == 5:
        print(f"[GSTR-3B] {purpose} calculation: Setting due day as 30th of next month for month: {return_month}, year: {financial_year}")
        day = 30
    elif financial_year == financial_year_2024_25 and parse_month(return_month) == 12:
        print(f"[GSTR-3B] {purpose} calculation: Setting due day as 22nd of next month for month: {return_month}, year: {financial_year}")
        day = 22
    else:
        print(f"[GSTR-3B] {purpose} calculation: Setting due day as 20th of next month for month: {return_month}, year: {financial_year}")
        day = 20
    return day