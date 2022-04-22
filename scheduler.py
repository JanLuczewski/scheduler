from openpyxl import load_workbook
from pathlib import Path
from docxtpl import DocxTemplate

base_dir = Path(__file__).parent
workbook = load_workbook(filename=base_dir / "odbiory.xlsx")
sheet = workbook.active

word_template_path = base_dir / "formatka_A35.docx"
doc = DocxTemplate(word_template_path)
output_dir = base_dir / "OUTPUT"
output_dir.mkdir(exist_ok=True)

for row in sheet.iter_rows(min_row=2,values_only=True):
    Inspection_ID = row[1]
    Day = row[3]
    Starting_time = row[4]
    Description = row[15]
    context = {
        "Inspection_ID" : Inspection_ID,
        "Description" : Description,
        "Starting_time" : Starting_time,
        "Day" : Day,
    }
    doc.render(context)
    doc.save(output_dir / "plik_wypluty.docx")


