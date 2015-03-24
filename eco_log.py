#!/usr/bin/env python3

# standard libraries
import pywintypes
import textwrap
from collections import defaultdict

# Third Party Open Source Libs
from xlwings import Workbook, Sheet, Range, RowCol   # Control Excel via COM. https://pypi.python.org/pypi/xlwings/0.3.4
import six                                   # https://pypi.python.org/pypi/six/1.9.0


def calc_first_blank():
    return len(Range('A1').vertical) + 1


def error_text(error):
    hr, msg, exc, arg = error.args
    return textwrap.fill(exc[2])


def save_workbook(workbook, filename):
    try:
        workbook.save(filename)
    except pywintypes.com_error as error:
        print(error_text(error))
        return False

    return True


def open_workbook(workbook_path):

    try:
        wb = Workbook(workbook_path)
    except pywintypes.com_error as error:
        print(error_text(error))
        return None

    return wb


def activate_sheet(sheet_name):
    try:
        Sheet(sheet_name).activate()
    except pywintypes.com_error as error:
        print(error_text(error))
        return None


def read_workbook(wb_path, sheet_name):

    wb = open_workbook(wb_path)

    if not wb:
        return 0

    activate_sheet(sheet_name)
    sheet_data = Range('A4').table.value

    eco_log_data = defaultdict()

    for row in sheet_data:
        if row[1]:
            eco_num = str(round(row[0]))
            eco_log_data[eco_num] = defaultdict()
            eco_log_data[eco_num]["date_assigned"] = row[1]
            eco_log_data[eco_num]["initiator"] = row[2]
            eco_log_data[eco_num]["part_number"] = row[3]
            eco_log_data[eco_num]["project_data"] = row[4]

    # if not save_workbook(wb, "c:\cid-tool\cid\output.xlsx"):
    #     return 0

    return 1


def main():
    read_workbook("c:\cid-tool\cid\ECO.xlsx", "ALL ECO's")

if __name__ == "__main__":
    main()

