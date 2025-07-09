from pathlib import Path
import xlwings as xw

def main():
    BASE_DIR = Path(__file__).resolve().parent
    INPUT_DIR = BASE_DIR / "Input"
    OUTPUT_DIR = BASE_DIR / "Output"
    
    # สร้างโฟลเดอร์ Output หากไม่มี
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    # หาทุกโฟลเดอร์ในโฟลเดอร์ Input
    project_folders = [f for f in INPUT_DIR.iterdir() if f.is_dir()]
    
    for project_folder in project_folders:
        print(f"Processing project: {project_folder.name}")
        
        # สร้างโฟลเดอร์ Output สำหรับโครงการนี้
        project_output_folder = OUTPUT_DIR / project_folder.name
        project_output_folder.mkdir(parents=True, exist_ok=True)
        
        # หาไฟล์ Excel ในโฟลเดอร์โครงการ
        excel_files = list(project_folder.glob("*.xlsx"))
        if not excel_files:
            print(f"Warning: No Excel files found in {project_folder.name}")
            continue
                       
        # ใช้ไฟล์แรกที่พบ
        excel_file = excel_files[0]
        print(len(str(excel_file)))

        with xw.App(visible=False) as app:
            wb = app.books.open(excel_file)
            for sheet in wb.sheets:
                wb_new = app.books.add()
                sheet.copy(after=wb_new.sheets[0])
                wb_new.sheets[0].delete()
                wb_new.save(project_output_folder / f"{sheet.name}.xlsx")
                wb_new.close()
            wb.close()
        
        print(f"✓ Completed A_Page_Separator for: {project_folder.name}")

if __name__ == "__main__":
    main()