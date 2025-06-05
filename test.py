# test_gstr3b.py
from utils.extractors.gstr3b import extract_table_3_1

df = extract_table_3_1("uploaded_files/22AAEFL1333K1ZO/GSTR-3B/GSTR3B_22AAEFL1333K1ZO_042023.pdf")
print(df)
