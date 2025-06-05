import pdfplumber
import pandas as pd

index_map = {
        "3.1": 2,
        "3.1.1": 3,
        "3.2": 4,
        # "4": 6,
        "5": 7,
        "5.1":8,
        "6.1": 9
}

def extract_fixed_tables_from_gstr3b(pdf_path):
    """Extract tables using fixed position assumptions (e.g., 4th table = 3.1)."""
    all_tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table and len(table) > 1 and len(table[0]) > 1:
                    df = pd.DataFrame(table)
                    df.columns = df.iloc[0]
                    df = df[1:].reset_index(drop=True)
                    all_tables.append(df)

    table_map = {}
    for key, idx in index_map.items():
        if idx < len(all_tables):
            table_map[key] = all_tables[idx]
    return table_map
