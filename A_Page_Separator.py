from pathlib import Path
import xlwings as xw

def main():
    FILE_NAME = "MASTER PLAN แผนขั้นตอนและผลการปฏิบัติงาน.xlsx"
    BASE_DIR = Path(__file__).parent
    EXCEL_FILE = BASE_DIR / f"Input/{FILE_NAME}"
    OUTPUT_DIR = BASE_DIR / "Output"

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with xw.App(visible=False) as app:
        wb = app.books.open(EXCEL_FILE)
        for sheet in wb.sheets:
            wb_new = app.books.add()
            sheet.copy(after=wb_new.sheets[0])
            wb_new.sheets[0].delete()
            wb_new.save(OUTPUT_DIR / f"{sheet.name}.xlsx")
            wb_new.close()

if __name__ == "__main__":
    main()
