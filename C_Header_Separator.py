import pandas as pd
import re
from pathlib import Path

def process_single_file(file_path, output_dir):
    """Process a single Excel file and separate headers"""
    # อ่านไฟล์ Excel โดยใช้ชื่อไฟล์ที่กำหนด
    FILE_NAME = "output.xlsx"
    INPUT_PATH = file_path if file_path.name == FILE_NAME else file_path / FILE_NAME
    
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
    BASE_OUTPUT = output_dir
    df_main.to_excel(BASE_OUTPUT / "main_data.xlsx", index=False)
    df_progress.to_excel(BASE_OUTPUT / "progress_tracker.xlsx", index=False)

def main():
    FILE_NAME = "output.xlsx"
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / "Output"
    
    # ตรวจสอบว่ามีโฟลเดอร์ Output หรือไม่
    if not OUTPUT_DIR.exists():
        print(f"Error: Output directory not found at {OUTPUT_DIR}")
        return
    
    # หาทุกโฟลเดอร์ในโฟลเดอร์ Output
    project_folders = [f for f in OUTPUT_DIR.iterdir() if f.is_dir()]
    
    # ถ้าไม่มีโฟลเดอร์ย่อย ให้ใช้ Output โฟลเดอร์เอง
    if not project_folders:
        print("No project folders found. Processing Output folder directly...")
        INPUT_PATH = OUTPUT_DIR / FILE_NAME
        if INPUT_PATH.exists():
            try:
                process_single_file(INPUT_PATH, OUTPUT_DIR)
                print("✓ Completed C_Header_Separator for: Output folder")
            except Exception as e:
                print(f"Error processing Output folder: {e}")
        else:
            print(f"Warning: {FILE_NAME} not found in Output folder")
        return
    
    # ประมวลผลแต่ละโฟลเดอร์
    for project_folder in project_folders:
        print(f"Processing project: {project_folder.name}")
        
        INPUT_PATH = project_folder / FILE_NAME
        if not INPUT_PATH.exists():
            print(f"Warning: {FILE_NAME} not found in {project_folder.name}")
            continue
        
        try:
            process_single_file(INPUT_PATH, project_folder)
            print(f"✓ Completed C_Header_Separator for: {project_folder.name}")
        except Exception as e:
            print(f"Error processing {project_folder.name}: {e}")

if __name__ == "__main__":
    main()