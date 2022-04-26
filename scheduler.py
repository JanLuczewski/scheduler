from openpyxl import load_workbook
from pathlib import Path
from docxtpl import DocxTemplate
import argparse
import shutil

""" scheduler.py służy do tworzenia raportów dla inspekcji.
Program wywołuje się wpisując w terminalu 
    python3 scheduler.py source template
przykładowo
    python3 scheduler.py odbiory.xlsx formatka_A35
wygenerowane raporty można znaleźć w folderze OUTPUT
"""

parser = argparse.ArgumentParser(description='wypluwa raporty dla odbiorow')
parser.add_argument('source', metavar='zrodlo', type=str, help='source   czyli źródłowy plik excel z odbiorami zgłoszonymi do Calypso')
parser.add_argument('template', metavar='formatka', type=str, help='template   czyli plik służący za formatkę dla danego projektu stoczniowego')
args = parser.parse_args()

source = args.source
template = args.template



base_dir = Path(__file__).parent
workbook = load_workbook(filename=base_dir / source)
sheet = workbook.active

word_template_path = base_dir / template
doc = DocxTemplate(word_template_path)
output_dir = base_dir / "OUTPUT"
output_dir.mkdir(exist_ok=True)


for row in sheet.iter_rows(min_row=2,values_only=True):
    Inspection_ID = row[1]
    Day = row[3]
    Starting_time = row[4]
    Description = row[15]
    Inspector = row[10]
    JanLuczewski = "JAN LUCZEWSKI"
    if Inspector == JanLuczewski:
        context = {
            "Inspection_ID" : Inspection_ID,
            "Description" : Description,
            "Starting_time" : Starting_time,
            "Day" : Day,
        }
        doc.render(context)
        doc.save(output_dir / f"{Inspection_ID}.docx")
shutil.make_archive(source,'zip','OUTPUT')



