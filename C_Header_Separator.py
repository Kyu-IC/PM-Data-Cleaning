import pandas as pd
import re
from pathlib import Path

# === Path ตั้งต้น ===
FILE_NAME = "output.xlsx"
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "Output"
INPUT_PATH = OUTPUT_DIR / FILE_NAME
print(INPUT_PATH)

# === โหลดไฟล์แบบ MultiIndex ===
df = pd.read_excel(INPUT_PATH, header=[0, 1])

# === แยก 3 คอลัมน์แรกเก็บไว้ ===
reference_columns = df.columns[:3]
df_reference = df[reference_columns]

# === ตรวจจับคอลัมน์วันที่ ===
date_columns = []
for col in df.columns:
    col_str = str(col[1])  # ใช้ header ชั้นล่าง
    try:
        parsed = pd.to_datetime(col_str, format="%d %b %y", dayfirst=True)
        date_columns.append(col)
    except:
        if re.search(r'\d{1,2}\s*(ม\.ค\.|ก\.พ\.|มี\.ค\.|เม\.ย\.|พ\.ค\.|มิ\.ย\.|ก\.ค\.|ส\.ค\.|ก\.ย\.|ต\.ค\.|พ\.ย\.|ธ\.ค\.)', col_str):
            date_columns.append(col)

# === แยกข้อมูลหลักกับ progress ===
main_columns = [col for col in df.columns if col not in date_columns and col not in reference_columns]
df_main = df[main_columns]
df_progress = df[date_columns]

# === FLATTEN MULTIINDEX ก่อน save ===
df_reference.columns = ['_'.join([str(i) for i in col]).strip() for col in df_reference.columns]
df_main.columns = ['_'.join([str(i) for i in col]).strip() for col in df_main.columns]
df_progress.columns = ['_'.join([str(i) for i in col]).strip() for col in df_progress.columns]

# === Save ไฟล์ ===
df_reference.to_excel(OUTPUT_DIR / "reference_data.xlsx", index=False)
df_main.to_excel(OUTPUT_DIR / "main_data.xlsx", index=False)
df_progress.to_excel(OUTPUT_DIR / "progress_tracker.xlsx", index=False)
