from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import date
import os

OUTPUT_FILE = "output/contabilita_SG_Impianti_2026.xlsx"

FOGLI = [
    "Prima Nota",
    "Clienti",
    "Fornitori",
    "Banca",
    "Cassa",
    "IVA Vendite",
    "IVA Acquisti",
    "Costi",
    "Ricavi",
    "Riepilogo"
]

def build_workbook(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = Workbook()

    # Rimuove il foglio di default
    wb.remove(wb.active)

    for nome in FOGLI:
        ws = wb.create_sheet(title=nome)
        ws["A1"] = f"Contabilità SG Impianti – 2026 – {nome}"
        ws["A1"].font = Font(size=14, bold=True)
        ws.append(["Data", "Voce", "Descrizione", "Entrate", "Uscite", "Note"])
        ws.freeze_panes = "A3"
        ws.append([date.today().isoformat(), "", "", "", "", ""])

sheetnames = wb.sheetnames
print("SHEETS_COUNT:", len(sheetnames))
print("SHEETS:", sheetnames)

if len(sheetnames) < 10:
    raise Exception(f"Workbook incorrect: only {len(sheetnames)} sheets")

    wb.save(path)

if __name__ == "__main__":
    build_workbook(OUTPUT_FILE)
