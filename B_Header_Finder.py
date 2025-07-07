import pandas as pd
from pathlib import Path

FILE_NAME = "แผนการปฎิบัติงาน (QTY).xlsx"
BASE_DIR = Path(__file__).resolve().parent
INPUT_PATH = BASE_DIR / "Output" / FILE_NAME

print(INPUT_PATH)


df_raw = pd.read_excel(INPUT_PATH)

header_row_idx = None
for i, row in df_raw.iterrows():
    if row.astype(str).str.contains("ลำดับ").any():
        header_row_idx = i
        break

if header_row_idx is not None:
    new_header = df_raw.iloc[header_row_idx]

    df = df_raw.iloc[header_row_idx + 1:].copy()
    df.columns = new_header.values
    df.reset_index(drop=True, inplace=True)

    print(df.head())

else:
    print("ไม่พบแถวที่มีคำว่า 'ลำดับ'")

df = df.dropna(axis=1, how='all')
df.to_excel(f"Output/output.xlsx", index=False)


