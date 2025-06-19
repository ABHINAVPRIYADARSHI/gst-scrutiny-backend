import pdfplumber
import pandas as pd
from tabulate import tabulate
table_header_rows_skip = {
    0: -1,
    1 : 4, # Table 4
    2 : 0, # Table 4
    3 : 4, # Table 5
    4 : -1, # Table 5
    # 5 : 5, not required
    # 6: 0,  not required
    # 7 : 3, # Table 6
    8 : 3,  # Table 7
    # 9 : 0,  Part of Table 7 not required
    10 : 3, # Table 8
    11 :3, # Table 9
}


def extract_fixed_tables_from_gstr9(pdf_path):
    print(f"Starting execution of function extract_fixed_tables_from_gstr9 for: {pdf_path}")
    all_tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                all_tables.append(table)

        for idx, skip_rows in table_header_rows_skip.items():
                if idx < len(all_tables):
                    df = pd.DataFrame(all_tables[idx])
                    df.columns = df.iloc[skip_rows]  # Set new header
                    df = df.drop(index=range(skip_rows+1))  # Drop the header rows
                    df = df.reset_index(drop=True)  # Reset index
                    all_tables[idx] = df
                    # print(f"Table no: {idx}")
                    # print(tabulate(df, tablefmt='grid', maxcolwidths=20))
    print(f"Starting execution of function extract_fixed_tables_from_gstr9 for: {pdf_path}")
