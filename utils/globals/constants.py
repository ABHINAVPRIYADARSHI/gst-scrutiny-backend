from datetime import date
import re

# Table position starts with Table header and not the first row of table.
OLD_TABLE_POSITIONS_GSTR_3B = {
    "1": {
        "start_row": 2,  # Excel row 3
        "end_row": 4,  # Excel row 5
        "start_col": 0,  # Column A
        "end_col": 1  # Column B
    },
    "2": {
        "start_row": 7,  # Excel row 8
        "end_row": 12,  # Excel row 13
        "start_col": 0,  # Column A
        "end_col": 1  # Column B
    },
    "3.1": {
        "start_row": 15,  # Excel row 16
        "end_row": 20,  # Excel row 21
        "start_col": 0,  # Column A
        "end_col": 5  # Column F
    },
    "3.2": {
        "start_row": 23,  # Excel row 24
        "end_row": 26,  # Excel row 27
        "start_col": 0,  # Column A
        "end_col": 2  # Column C
    },
    "4": {
        "start_row": 29,  # Excel row 30
        "end_row": 42,  # Excel row 43
        "start_col": 0,  # Column A
        "end_col": 4  # Column E
    },
    "5": {
        "start_row": 46,  # Excel row 47
        "end_row": 48,  # Excel row 49
        "start_col": 0,  # Column A
        "end_col": 2  # Column C
    },
    "5.1": {
        "start_row": 52,  # Excel row 53
        "end_row": 55,  # Excel row 56
        "start_col": 0,  # Column A
        "end_col": 4  # Column E
    },
    "6.1": {
        "start_row": 58,  # Excel row 59
        "end_row": 68,  # Excel row 69
        "start_col": 0,  # Column A
        "end_col": 8  # Column I
    },
}
NEW_TABLE_POSITIONS_GSTR_3B = {
    "1": {
        "start_row": 2,  # Excel row 3
        "end_row": 4,  # Excel row 5
        "start_col": 0,  # Column A
        "end_col": 1  # Column B
    },
    "2": {
        "start_row": 7,  # Excel row 8
        "end_row": 12,  # Excel row 13
        "start_col": 0,  # Column A
        "end_col": 1  # Column B
    },
    "3.1": {
        "start_row": 15,  # Excel row 16
        "end_row": 20,  # Excel row 21
        "start_col": 0,  # Column A
        "end_col": 5  # Column F
    },
    "3.1.1": {
        "start_row": 23,  # Excel row 24
        "end_row": 25,  # Excel row 26
        "start_col": 0,  # Column A
        "end_col": 5  # Column F
    },
    "3.2": {
        "start_row": 28,  # Excel row 29
        "end_row": 31,  # Excel row 32
        "start_col": 0,  # Column A
        "end_col": 2  # Column C
    },
    "4": {
        "start_row": 34,  # Excel row 35
        "end_row": 47,  # Excel row 48
        "start_col": 0,  # Column A
        "end_col": 4  # Column E
    },
    "5": {
        "start_row": 51,  # Excel row 52
        "end_row": 53,  # Excel row 54
        "start_col": 0,  # Column A
        "end_col": 2  # Column C
    },
    "5.1": {
        "start_row": 56,  # Excel row 57
        "end_row": 59,  # Excel row 60
        "start_col": 0,  # Column A
        "end_col": 4  # Column E
    },
    "6.1": {
        "start_row": 62,  # Excel row 63
        "end_row": 72,  # Excel row 73
        "start_col": 0,  # Column A
        "end_col": 8  # Column I
    },
}

late_fee_headers = [
    "Financial Year", "Return Month", "Date of ARN", "Due Date", "Days Late",
    "Late fee @ ₹100/per day", "Late fee applicable"
]

oldFormat = "OLD_FORMAT"
newFormat = "NEW_FORMAT"
b2b_merged_sheet = "B2B_merged"
cdnr_merged_sheet = "CDNR_merged"
debit_note = "DEBIT NOTE"
credit_note = "CREDIT NOTE"
isd_merged_sheet = "ISD_merged"
impg_merged_sheet = "IMPG_merged"
impg_sez_merged_sheet = "IMPG SEZ_merged"
tcs_merged_sheet = "TCS_merged"
tds_merged_sheet = "TDS_merged"
eco_merged_sheet = "ECO_merged"
diffInRCM_ITC = "diffInRCM_ITC"
estimatedITCReversal = "estimatedITCReversal"
diffInRCMPayment = "diffInRCMPayment"
ewb_in_MIS_report = "EWB_in_MIS_report_merged"
ewb_out_MIS_report = "EWB_Out_MIS_report_merged"
total_string = "-Total"
string_Y = "Y"
string_YES = "YES"
string_yes = "Yes"
string_no = "No"
string_NO = "NO"
string_N = "N"
sheet_overview = "Overview"
empty_string = ''
int_zero = 0
int_one = 1
int_nine = 9
int_ten = 10
int_eighteen = 18
int_twenty = 20
int_twenty_one = 21
floating_zero = 0.00
string_zero = "0.00"
str_one = "1"
str_two = "2"
str_three_point_one = "3.1"
str_four = "4"
str_three_point_one_point_one = "3.1.1"
str_six_point_one = "6.1"
gstr1_analysis_dict = "gstr1_analysis_dict",
gstr3b_analysis_dict = "gstr3b_analysis_dict"
gstr3b_merged_dict = "gstr3b_merged_dict"
ewb_in_analysis_dict = "ewb_in_analysis_dict",
ewb_out_analysis_dict = "ewb_out_analysis_dict",
bo_comparison_summary_dict = "bo_comparison_summary_dict"
result_point_1 = "result_point_1"
result_point_2 = "result_point_2"
result_point_3 = "result_point_3"
result_point_4 = "result_point_4"
result_point_5 = "result_point_5"
result_point_6 = "result_point_6"
result_point_7 = "result_point_7"
result_point_8 = "result_point_8"
result_point_9 = "result_point_9"
result_point_10 = "result_point_10"
result_point_11 = "result_point_11"
result_point_12 = "result_point_12"
result_point_13 = "result_point_13"
result_point_14 = "result_point_14"
result_point_15 = "result_point_15"
result_point_16 = "result_point_16"
result_point_17 = "result_point_17"
result_point_18 = "result_point_18"
result_point_19 = "result_point_19"
result_point_20 = "result_point_20"
result_point_21 = "result_point_21"
result_point_22 = "result_point_22"

financial_year_2022_23 = "2022-23"
financial_year_2023_24 = "2023-24"
financial_year_2024_25 = "2024-25"


def extract_table_with_header(df, table_key, table_positions):
    pos = table_positions[table_key]
    raw_table = df.iloc[
                pos["start_row"]: pos["end_row"] + 1,
                pos["start_col"]: pos["end_col"] + 1
                ]
    # Use first row as header
    raw_table.columns = raw_table.iloc[0]
    cleaned_table = raw_table.drop(raw_table.index[0]).reset_index(drop=True)
    return cleaned_table


def convert_to_number(value):
    try:
        # Clean value: remove commas, strip spaces
        cleaned = str(value).replace(',', '').strip()
        # Convert to float (handles int too)
        return float(cleaned)
    except (ValueError, TypeError):
        return value  # Leave as-is if not convertible


month_lookup = {
    'April': 4, 'Apr': 4, 'May': 5,
    'June': 6, 'Jun': 6, 'July': 7, 'Jul': 7,
    'August': 8, 'Aug': 8, 'September': 9, 'Sep': 9,
    'October': 10, 'Oct': 10, 'November': 11, 'Nov': 11,
    'December': 12, 'Dec': 12, 'January': 1, 'Jan': 1,
    'February': 2, 'Feb': 2, 'March': 3, 'Mar': 3
}


# Convert "April", "May", etc. and "2021-22" to datetime
def parse_month_year(month_name, financial_year):
    month_num = parse_month(month_name)
    fy_start_year = int(financial_year.split('-')[0])
    year = fy_start_year if month_num >= 4 else fy_start_year + 1
    return date(year, month_num, 1)


def parse_month(month_name):
    if '-' in month_name:  # If month_name has '-', take second part
        month_name = month_name.split('-')[-1].strip()
    return month_lookup[month_name]


def clean_and_parse_number(text):
    # Remove special characters like newlines, non-breaking spaces, tabs, etc.
    cleaned = re.sub(r'[^\d.-]+', '', str(text))
    try:
        return float(cleaned)
    except ValueError:
        return 0
