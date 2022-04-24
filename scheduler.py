from openpyxl import load_workbook
from pathlib import Path
from docxtpl import DocxTemplate
import argparse

parser = argparse.ArgumentParser(description='wypluwa raporty dla odbiorow')
parser.add_argument('source', metavar='zrodlo', type=str, help='wpisz nazwe pliku zrodlowego')
args = parser.parse_args()

source = args.source

base_dir = Path(__file__).parent
workbook = load_workbook(filename=base_dir / source)
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
    doc.save(output_dir / f"{Inspection_ID}.docx")


