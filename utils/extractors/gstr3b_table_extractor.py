import pdfplumber
import pandas as pd
from tabulate import tabulate

from utils.globals.constants import int_nine, int_zero, int_one, int_ten

index_map_format_before_2022 = {
    "1": 0,
    "2": 1,
    "3.1": 2,
    "3.2": 3,
    "4": 4,
    "5": 5,
    "5.1": 6,
    "6.1": 7,
    "7": 8
}
index_map_format_later_than_2022 = {
    "1": 0,
    "2": 1,
    "3.1": 2,
    "3.1.1": 3,
    "3.2": 4,
    "4": 6,  # Table 4 Is added manually later on since its concatenation of two tables 5 & 6
    "5": 7,  # There seems to be some issue with 5 & 5.1. While creating GSTR-3B merged,
    "5.1": 8,  # the table header is not populated properly. We are proceeding with this known
    "6.1": 9,  # defect as of now since we don't use tables 5 & 5.1 . Later on, it needs a fix.
    "7": 10
}


# This function receives one PDF file at a time and extracts all tables in it.
def extract_fixed_tables_from_gstr3b(pdf_path):
    print(f"Starting execution of function extract_fixed_tables_from_gstr3b for: {pdf_path}")
    all_tables = []
    table_map = {}
    """Extract tables using fixed position assumptions (e.g., 4th table = 3.1)."""
    print(f"")
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for i, table in enumerate(tables):
                if table and len(table) > 1 and len(table[0]) > 1:
                    df = pd.DataFrame(table)
                    if i not in [int_zero, int_one]:
                        df.columns = df.iloc[0]
                        df = df[1:].reset_index(drop=True)
                    all_tables.append(df)
                    # print(f"{df}")
                    # print("*****************")

    if len(all_tables) == int_nine:  # New format has 11 tables
        print(
            f"GSTR-3B file {pdf_path} uploaded is old format (earlier than or 2021-22) since it has {len(all_tables)} tables.")
        return extract_old_format_tables(all_tables, table_map)
    elif len(all_tables) == int_ten:  # New format has 10 tables (Sep 2024 onwards)
        print(
            f"GSTR-3B file {pdf_path} uploaded is new format (Sep 2024 onwards) since it has {len(all_tables)} tables.")
        return extract_new_format_tables(all_tables, table_map)
    else:
        print(
            f"GSTR-3B file {pdf_path} uploaded is new format (later than FY 2021-22) since it has {len(all_tables)} tables.")
        return extract_new_format_tables(all_tables, table_map)


def extract_old_format_tables(all_tables, table_map):
    for key, idx in index_map_format_before_2022.items():
        if idx < len(all_tables):
            table_map[key] = all_tables[idx]
    return table_map


def extract_new_format_tables(all_tables, table_map):
    for key, idx in index_map_format_later_than_2022.items():
        if idx < len(all_tables):
            if key == "4":
                processtable4(all_tables, table_map)
            else:
                table_map[key] = all_tables[idx]
    return table_map


def processtable4(all_tables, table_map):
    # manually add PDF table 4 to table_map since its the concatenation of two split
    # tables 5 & 6 present on 1st page end & 2nd page beginning of any GSTR_3B pdf
    # print(tabulate(all_tables[5], tablefmt='grid', maxcolwidths=20))
    # print("===================================")
    # print(tabulate(all_tables[6], tablefmt='grid', maxcolwidths=20))

    df5 = all_tables[5].copy()
    df6 = all_tables[6].copy()
    df6.columns = df5.columns
    df_merged = pd.concat([df5, df6], ignore_index=True)
    # print(tabulate(df_merged, tablefmt='grid', maxcolwidths=20))
    table_map["4"] = df_merged
