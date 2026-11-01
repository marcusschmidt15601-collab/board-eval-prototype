"""make_variants.py - build two broken-input test files from synthetic/raw.xlsx.
Contains NO real data.
  raw_zero.xlsx      : one zero answer on a 1-5 item (should be treated as missing)
  raw_yes_blank.xlsx : 'yes' to written rules, conditional question left blank (should be flagged)
"""
import openpyxl

src = "synthetic/raw.xlsx"

# 1. Zero on a 1-5 item: respondent 2 (Excel row 4), column 14
wb = openpyxl.load_workbook(src); ws = wb.active
ws.cell(row=4, column=14).value = 0
wb.save("synthetic/raw_zero.xlsx")

# 2. 'Yes' to written rules but conditional question blank: respondent 4 (row 6), columns 17 and 18
wb = openpyxl.load_workbook(src); ws = wb.active
ws.cell(row=6, column=17).value = 1
ws.cell(row=6, column=18).value = None
wb.save("synthetic/raw_yes_blank.xlsx")
print("saved raw_zero.xlsx and raw_yes_blank.xlsx")
