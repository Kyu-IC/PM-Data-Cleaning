
import pandas as pd
import re
from pathlib import Path

def main():
    FILE_NAME = "output.xlsx"
    BASE_DIR = Path(__file__).resolve().parent
    INPUT_PATH = BASE_DIR / "Output" / FILE_NAME

    df = pd.read_excel(INPUT_PATH, header=[0, 1])

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
    main_columns = [col for col in df.columns if col not in date_columns]
    df_main = df[main_columns]
    df_progress = df[date_columns]

    # === แทรกคอลัมน์อ้างอิง ===
    ref_cols = [
        ('รายละเอียด', 'เป้าหมาย'),
        ('รายละเอียด', 'ผลการปฏิบัติ'),
        ('รายละเอียด', 'ปริมาณผลการปฎิบัติ')
    ]

    for col in ref_cols:
        if col in df_main.columns:
            df_progress.insert(0, col[1], df[col])

    # === FLATTEN MULTIINDEX ก่อน save ===
    df_main.columns = ['_'.join([str(i) for i in col]).strip() for col in df_main.columns]
    df_progress.columns = ['_'.join([str(i) for i in col]).strip() for col in df_progress.columns]

    # === Save ===
    BASE_OUTPUT = BASE_DIR / "Output"
    df_main.to_excel(BASE_OUTPUT / "main_data.xlsx", index=False)
    df_progress.to_excel(BASE_OUTPUT / "progress_tracker.xlsx", index=False)

if __name__ == "__main__":
    main()