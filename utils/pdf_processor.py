import pdfplumber
import pandas as pd
import os

def process_pdf_files(file_paths: list[str], return_type: str) -> str:
    combined_data = []

    for path in file_paths:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    combined_data.append(df)

    master_df = pd.concat(combined_data).fillna(0)
    master_path = f"reports/{return_type}_master.csv"
    os.makedirs("reports", exist_ok=True)
    master_df.to_csv(master_path, index=False)
    return master_path
