#!/usr/bin/env python3

# standard libraries
import pywintypes
import textwrap
import sys
from collections import defaultdict

# Third Party Open Source Libs
from xlwings import Workbook, Sheet, Range  # Control Excel via COM. https://pypi.python.org/pypi/xlwings/0.3.4
import openpyxl                             # https://pypi.python.org/pypi/openpyxl/2.2.0

import cid


def calc_first_blank_xlwings():
    return len(Range('A1').vertical) + 1


def error_text(error):
    hr, msg, exc, arg = error.args
    return textwrap.fill(exc[2])


def save_workbook_xlwings(workbook, filename):
    try:
        workbook.save(filename)
    except pywintypes.com_error as error:
        print(error_text(error))
        return False

    return True


def open_workbook_xlwings(workbook_path):

    try:
        wb = Workbook(workbook_path)
    except pywintypes.com_error as error:
        print(error_text(error))
        return None

    return wb


def activate_sheet_xlwings(sheet_name):
    try:
        Sheet(sheet_name).activate()
    except pywintypes.com_error as error:
        print(error_text(error))
        return None


def read_openpyxl(wb_path, sheet_name, sheet_id):

    try:
        # openpyxl is a library for reading/writing Excel files.
        wb = openpyxl.load_workbook(wb_path, data_only=True)
    except openpyxl.utils.exceptions.InvalidFileException:
        cid.err_col('\n{}: Could not open {} at:'.format(sheet_id.upper(), sheet_id) +
                    '\n       {}'.format(wb_path))
        cid.exit_app()

    wb_sheet = wb.worksheets[0]
    try:
        wb_rows = wb_sheet.rows
    except AttributeError:
        cid.err_col('\n{}: No {} tab in {} at path:'.format(sheet_id.upper(), sheet_name, sheet_id) +
                    '\n\n     {}'.format(wb_path))
        sys.exit()

    eco_data = defaultdict()

    wb_update_table = []

    row_num = 1

    # marked_unused = []
    marked_unused = ["cancelled", "unused", "void", "reuse me!"]

    for row in wb_rows:
        if row_num < 4:
            row_num += 1
            continue

        if row[1].value and str(row[1].value).lower() not in marked_unused:
            eco_num = row[0].value
            print(eco_num)
            eco_data[eco_num] = defaultdict()
            eco_data[eco_num]["date_assigned"] = row[1].value
            eco_data[eco_num]["initiator"] = row[2].value
            eco_data[eco_num]["part_number"] = row[3].value
            eco_data[eco_num]["project_data"] = row[4].value
            eco_data[eco_num]["incorp_date"] = row[5].value
            wb_update_table.append([eco_num, row[1].value, row[2].value, row[3].value, row[4].value, row[5].value])

        row_num += 1

    return eco_data


def main():
    eco_data = read_openpyxl(r"c:\cid-tool\cid\ECO.xlsx", "ALL ECO's", "ECO Log")
    if eco_data:
        print("Completed.")

if __name__ == "__main__":
    main()

