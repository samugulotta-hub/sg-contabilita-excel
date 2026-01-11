# sg-contabilita-excel

## Requisiti
- Python 3.11+

## Istruzioni Windows

1. Apri PowerShell nella cartella del progetto.
2. Crea e attiva un virtualenv:
   ```powershell
   py -3.11 -m venv .venv
   .venv\Scripts\Activate.ps1
   ```
3. Installa le dipendenze:
   ```powershell
   pip install -r requirements.txt
   ```
4. Genera l'Excel:
   ```powershell
   python tools\build_workbook.py
   ```

Il file generato si trova in:
```
output\contabilita_SG_Impianti_2026.xlsx
```
