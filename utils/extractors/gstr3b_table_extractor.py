import pdfplumber
import pandas as pd

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
        "4": 6, # Table 4 Is added manually later on since its concatenation of two tables 5 & 6
        "5": 7,
        "5.1": 8,
        "6.1": 9,
        "7": 10
}

# This function receives one PDF file at a time and extracts all tables in it.
def extract_fixed_tables_from_gstr3b(pdf_path):
    all_tables = []
    table_map = {}
    """Extract tables using fixed position assumptions (e.g., 4th table = 3.1)."""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table and len(table) > 1 and len(table[0]) > 1:
                    df = pd.DataFrame(table)
                    df.columns = df.iloc[0]
                    df = df[1:].reset_index(drop=True)
                    all_tables.append(df)
    print(len(all_tables))
    if len(all_tables) == 9: # New format has 11 tables
        print("GSTR-3B file uploaded is old format (earlier than 2022.)")
        return extract_old_format_tables(all_tables, table_map)
    else:
        print("GSTR-3B file uploaded is new format (older than 2022.)")
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
                processTable4(all_tables, table_map)
            else:
                table_map[key] = all_tables[idx]
    return table_map


def processTable4(all_tables, table_map):
    # manually add PDF table 4 to table_map since its the concatenation of two split
    # tables 5 & 6 present on 1st page end & 2nd page beginning of any GSTR_3B pdf
    df6 = all_tables[6].copy() #temp copy of table[6]
    original_header_row = list(df6.columns)  # Save the wrongly assigned header
    df6.columns = range(len(df6.columns))  # Temporarily remove headers
    df6 = pd.concat([pd.DataFrame([original_header_row]), df6], ignore_index=True)
    df6.columns = all_tables[5].columns # Now assign the correct header from df5
    table_map["4"] = pd.concat([all_tables[5], df6], ignore_index=True) # Combine both tables
    # print(table_map)  # The print statement can print all tables gathered from all pages of a a given pdf file.
