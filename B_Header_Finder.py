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
    # แยก metadata และลบคอลัมน์ Unnamed
    df_metadata = df_raw.iloc[:header_row_idx].copy()
    df_metadata = df_metadata.loc[:, ~df_metadata.columns.str.contains("^Unnamed")]
    df_metadata.reset_index(drop=True, inplace=True)
    df_metadata.to_excel(BASE_DIR / "Output" / "metadata.xlsx", index=False)

    # จัดการส่วนข้อมูลหลัก
    new_header = df_raw.iloc[header_row_idx]
    df = df_raw.iloc[header_row_idx + 1:].copy()
    df.columns = new_header.values
    df.reset_index(drop=True, inplace=True)
    df = df.dropna(axis=1, how='all')

    df.to_excel(BASE_DIR / "Output" / "output.xlsx", index=False)

    print("✅ แยก metadata และข้อมูลหลักเรียบร้อยแล้ว")
else:
    print("❌ ไม่พบแถวที่มีคำว่า 'ลำดับ'")
