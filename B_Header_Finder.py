import pandas as pd
from pathlib import Path

def main():
    BASE_DIR = Path(__file__).resolve().parent
    OUTPUT_DIR = BASE_DIR / "Output"
    
    # หาทุกโฟลเดอร์ในโฟลเดอร์ Output
    project_folders = [f for f in OUTPUT_DIR.iterdir() if f.is_dir()]
    
    for project_folder in project_folders:
        print(f"Processing project: {project_folder.name}")
        
        # หาไฟล์ที่มีคำว่า "แผนการปฎิบัติงาน (QTY)" หรือคล้ายๆ
        possible_files = list(project_folder.glob("*แผนการปฎิบัติงาน*.xlsx"))
        if not possible_files:
            possible_files = list(project_folder.glob("*QTY*.xlsx"))
        if not possible_files:
            print(f"Warning: No suitable file found for {project_folder.name}")
            continue
        
        input_file = possible_files[0]
        df_raw = pd.read_excel(input_file)

        header_row_idx = None
        for i, row in df_raw.iterrows():
            if row.astype(str).str.contains("ลำดับ").any():
                header_row_idx = i
                break

        if header_row_idx is not None:
            # จัดการส่วนข้อมูลหลัก (ข้ามส่วน metadata)
            new_header = df_raw.iloc[header_row_idx]
            df = df_raw.iloc[header_row_idx + 1:].copy()
            df.columns = new_header.values
            df.reset_index(drop=True, inplace=True)
            
            # ลบคอลัมน์ Unnamed และคอลัมน์ว่าง
            df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
            df = df.dropna(axis=1, how='all')

            df.to_excel(project_folder / "output.xlsx", index=False)
            print(f"✓ Completed B_Header_Finder for: {project_folder.name}")
        else:
            print(f"Warning: Header row not found in {project_folder.name}")

if __name__ == "__main__":
    main()