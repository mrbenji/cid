import bdt_utils
from openpyxl import load_workbook

slog = load_workbook('11629.xlsm')
pn_sheet = slog.get_sheet_by_name('PS1')
pn_rows = pn_sheet.rows
row_num = 0
pn_table = []
for row in pn_rows:
    row_num += 1
    for cell in row:
        if row_num < 5 or not pn_sheet['A'+str(cell.row)].value or cell.column not in "ABCFG":
            continue

        if cell.column == "A":
            pn_table.append({"pn":cell.value, "rev":None, "desc":None, "indent":None})

        if cell.column == "B":
            pn_table[-1]["rev"] = cell.value

        if cell.column == "C":
            if cell.value:
                pn_table[-1]["rev"] = cell.value

        if cell.column == "F":
            if cell.value:
                pn_table[-1]["desc"] = cell.value
                pn_table[-1]["indent"] = int("{:.0f}".format(cell.style.alignment.indent))

for part in pn_table:
    print "{}{} Rev. {}  -  {}".format("  " * part["indent"], part["pn"], part["rev"], part["desc"])
