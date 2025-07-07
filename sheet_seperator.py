from pathlib import Path
import xlwings as xw

BASE_DIR = Path(__file__).parent
EXCEL_FILE = BASE_DIR / "MASTER PLAN แผนขั้นตอนและผลการปฏิบัติงาน.xlsx"
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

