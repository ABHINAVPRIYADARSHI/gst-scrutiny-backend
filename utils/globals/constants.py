from datetime import date
import re

OLD_TABLE_POSITIONS_GSTR_3B = {
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
    "3.1": {
        "start_row": 13,  # Excel row 14
        "end_row": 18,  # Excel row 19
        "start_col": 0,  # Column A
        "end_col": 5  # Column F
    },
    "3.1.1": {
        "start_row": 21,  # Excel row 22
        "end_row": 23,  # Excel row 24
        "start_col": 0,  # Column A
        "end_col": 5  # Column F
    },
    "3.2": {
        "start_row": 26,  # Excel row 27
        "end_row": 29,  # Excel row 30
        "start_col": 0,  # Column A
        "end_col": 2  # Column C
    },
    "4": {
        "start_row": 32,  # Excel row 33
        "end_row": 45,  # Excel row 46
        "start_col": 0,  # Column A
        "end_col": 4  # Column E
    },
    "5": {
        "start_row":48,  # Excel row 49
        "end_row": 50,  # Excel row 51
        "start_col": 0,  # Column A
        "end_col": 2  # Column C
    },
    "5.1": {
        "start_row": 53,  # Excel row 54
        "end_row": 56,  # Excel row 57
        "start_col": 0,  # Column A
        "end_col": 4  # Column E
    },
    "6.1": {
        "start_row": 59,  # Excel row 60
        "end_row": 69,  # Excel row 70
        "start_col": 0,  # Column A
        "end_col": 8  # Column I
    },
}
oldFormat = "OLD_FORMAT"
newFormat = "NEW_FORMAT"
diffInRCM_ITC = "diffInRCM_ITC"
estimatedITCReversal = "estimatedITCReversal"
diffInRCMPayment = "diffInRCMPayment"
ewb_in_MIS_report = "EWB_in_MIS_report_merged"
ewb_out_MIS_report = "EWB_Out_MIS_report_merged"
int_zero = 0
int_one = 1
int_nine = 9
int_eighteen = 18
floating_zero = 0.00
str_one = "1"
str_two = "2"
str_three_point_one_point_one = "3.1.1"
str_six_point_one = "6.1"
gstr1 = "GSTR-1",
gstr2a = "GSTR-2A",
gstr3b = "GSTR-3B"
gstr9 = "GSTR-9",
ewbIn = "EWB-IN",
ewbOut = "EWB-OUT",
bo_comparison = "BO_comparison"
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


def parse_month_year(month_name, financial_year):
    # Convert "April", "May", etc. and "2021-22" to datetime
    month_lookup = {
        'April': 4, 'Apr': 4, 'May': 5,
        'June': 6, 'Jun': 6, 'July': 7, 'Jul': 7,
        'August': 8, 'Aug': 8, 'September': 9, 'Sep': 9,
        'October': 10, 'Oct': 10, 'November': 11, 'Nov': 11,
        'December': 12, 'Dec': 12, 'January': 1, 'Jan': 1,
        'February': 2, 'Feb': 2, 'March': 3, 'Mar': 3
    }
    if '-' in month_name:      # If month_name has '-', take second part
        month_name = month_name.split('-')[-1].strip()
    month_num = month_lookup[month_name]
    fy_start_year = int(financial_year.split('-')[0])
    year = fy_start_year if month_num >= 4 else fy_start_year + 1
    return date(year, month_num, 1)


def clean_and_parse_number(text):
    # Remove special characters like newlines, non-breaking spaces, tabs, etc.
    cleaned = re.sub(r'[^\d.-]+', '', str(text))
    try:
        return float(cleaned)
    except ValueError:
        return 0
