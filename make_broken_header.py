"""make_broken_header.py - copy synthetic/raw.xlsx and blank out ONE peer-rating header.
Contains NO real data. Tests whether calc.py fails loudly when a header is missing."""
import openpyxl

SRC, DST = "synthetic/raw.xlsx", "synthetic/raw_broken_header.xlsx"
PEER_START = 135          # first peer column, 0-based, as in generate_synthetic.py
MISSING_OFFSET = 3        # remove the 4th member's header in the first block (Jo Demo)

wb = openpyxl.load_workbook(SRC)
ws = wb.active
excel_col = PEER_START + MISSING_OFFSET + 1      # openpyxl is 1-based
removed = ws.cell(row=1, column=excel_col).value
ws.cell(row=1, column=excel_col).value = None
wb.save(DST)
print(f"removed header '{removed}' from column {excel_col}; saved {DST}")
