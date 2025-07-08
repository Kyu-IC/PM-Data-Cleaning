import pandas as pd
from pathlib import Path

def main():
    BASE_DIR = Path(__file__).resolve().parent
    file_path = BASE_DIR / "Output/main_data.xlsx"
    df_raw = pd.read_excel(file_path)

    # ---------- STEP 2: แยกแถวเป้าหมายและผลการปฏิบัติ ----------
    mask_target = df_raw["รายละเอียดของงาน_Unnamed: 2_level_1"] == "เป้าหมาย"
    mask_actual = df_raw["รายละเอียดของงาน_Unnamed: 2_level_1"] == "ผลการปฏิบัติ"

    df_target = df_raw[mask_target].reset_index(drop=True)
    df_actual = df_raw[mask_actual].reset_index(drop=True)

    # ---------- STEP 3: รีเนมคอลัมน์ผลการปฏิบัติ ----------
    df_actual = df_actual.rename(columns={
        "กำหนดการ_วันเริ่มต้น": "ผลการปฏิบัติ_วันเริ่มต้น",
        "กำหนดการ_วันสิ้นสุด": "ผลการปฏิบัติ_วันสิ้นสุด"
    })

    # ---------- STEP 4: รวมข้อมูลแนวนอน ----------
    df_combined = pd.concat([
        df_target.reset_index(drop=True),
        df_actual[["ผลการปฏิบัติ_วันเริ่มต้น", "ผลการปฏิบัติ_วันสิ้นสุด"]].reset_index(drop=True)
    ], axis=1)

    # ---------- STEP 5: ย้ายคอลัมน์ผลการปฏิบัติให้อยู่หลังวันสิ้นสุดของเป้าหมาย ----------
    insert_after = df_combined.columns.get_loc("กำหนดการ_วันสิ้นสุด")
    cols = df_combined.columns.tolist()
    cols_to_move = ["ผลการปฏิบัติ_วันเริ่มต้น", "ผลการปฏิบัติ_วันสิ้นสุด"]

    for col in cols_to_move:
        cols.remove(col)
    for col in reversed(cols_to_move):
        cols.insert(insert_after + 1, col)

    df_reordered = df_combined[cols]

    # ---------- STEP 6: ลบเฉพาะคอลัมน์ที่เป็น 'เป้าหมาย/ผลการปฏิบัติ' ----------
    col_to_drop = "รายละเอียดของงาน_Unnamed: 2_level_1"
    if col_to_drop in df_reordered.columns:
        df_reordered.drop(columns=[col_to_drop], inplace=True)

    # ---------- STEP 7: ล้างชื่อคอลัมน์ Unnamed ให้อ่านง่าย ----------
    df_reordered.columns = [
        str(col).replace("_Unnamed: 0_level_1", "")
                 .replace("_Unnamed: 1_level_1", "")
                 .replace("_Unnamed: 2_level_1", "")
        if not str(col).startswith("Unnamed") else ""
        for col in df_reordered.columns
    ]

    # ---------- STEP 8: เซฟไฟล์สุดท้าย ----------
    df_reordered.to_excel(BASE_DIR / "Output/Final_Clean.xlsx", index=False)

if __name__ == "__main__":
    main()