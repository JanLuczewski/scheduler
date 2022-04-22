from pathlib import  Path
import pandas as pd
from docxtpl import DocxTemplate

base_dir = Path(__file__).parent
word_template_path = base_dir / "formatka_A35.docx"
excel_path = base_dir / "odbiory.xlsx"
output_dir = base_dir / "OUTPUT"

""" Stwórz folder zrzutu dla dokumentów"""
output_dir.mkdir(exist_ok=True)

"""przerób informacje w excelu na DataFrame"""
df = pd.read_excel(excel_path, sheet_name="Data")


"""Przeiterój każdy wiersz po kolei i wyrenderój plik word, wybór record spowoduje wyplucie listy słowników 
, potem przeiterujesz tą liste"""
for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template_path)
    doc.render(record)
    output_path = output_dir / f"{record['Inspection_ID']}-{record['Day']}.docx"
