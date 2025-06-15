import pandas as pd
import os


def process_csv_files(file_paths: list[str], return_type: str) -> str:
    df_list = [pd.read_csv(path) for path in file_paths]
    master_df = pd.concat(df_list).fillna(0)
    master_path = f"reports/{return_type}_master.csv"
    os.makedirs("reports", exist_ok=True)
    master_df.to_csv(master_path, index=False)
    return master_path
