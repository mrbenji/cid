#!/usr/bin/env python3

ECO_LOG_PATH = r"c:\cid-tool\cid\ECO.xlsx"

# ECO_LOG_PATH = \
#     r"\\us.ray.com\SAS\ast\eng\Operations\cm\Internal\Doc_Ctrl\hcm\DOC_CTRL_GENERAL\LOGS\eco_trackinglog\ECO.xlsx"

# standard libraries
import pywintypes
import textwrap
import sys

# Third Party Open Source Libs
from xlwings import Workbook, Sheet, Range  # Control Excel via COM. https://pypi.python.org/pypi/xlwings/0.3.4
import openpyxl                             # https://pypi.python.org/pypi/openpyxl/2.2.0

import cid
from cid_classes import Part


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


def eco_log_formulas_to_numbers(wb_path, sheet_name):

    wb = open_workbook_xlwings(wb_path)

    if not wb:
        return 0

    activate_sheet_xlwings(sheet_name)

    col_to_update = Range('A4:A{}'.format(calc_first_blank_xlwings())).value

    for row_num in range(0, len(col_to_update)):
        col_to_update[row_num] = round(col_to_update[row_num])

    Range('A4:E{}'.format(calc_first_blank_xlwings())).value = col_to_update

    if not save_workbook_xlwings(wb, ECO_LOG_PATH):
        return 0

    return 1


def update_eco_log(wb_path, sheet_name, row_num, part):

    wb = open_workbook_xlwings(wb_path)

    if not wb:
        return 0

    activate_sheet_xlwings(sheet_name)
    row_to_update = Range('A{r}:E{r}'.format(r=row_num)).value
    row_to_update[3] = "{} Rev. {}".format(part.number, part.max_rev.name)
    Range('A{r}:E{r}'.format(r=row_num)).value = row_to_update

    if not save_workbook_xlwings(wb, ECO_LOG_PATH):
        return 0

    return 1


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

    eco_log_data = {}

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
            eco_log_data[eco_num] = {}
            eco_log_data[eco_num]["row_num"] = row_num
            eco_log_data[eco_num]["date_assigned"] = row[1].value
            eco_log_data[eco_num]["initiator"] = row[2].value
            eco_log_data[eco_num]["part_number"] = row[3].value
            eco_log_data[eco_num]["project_data"] = row[4].value
            eco_log_data[eco_num]["incorp_date"] = row[5].value

        row_num += 1

    return eco_log_data


def main():
    eco_log_data = read_openpyxl(ECO_LOG_PATH, "ALL ECO's", "ECO Log")

    # if eco_log_data:
    #     eco_log_formulas_to_numbers(ECO_LOG_PATH, "ALL ECO's")
    test_part = Part("123-123456-01")
    test_part.add_rev("C")

    if eco_log_data:
        update_eco_log(ECO_LOG_PATH, "ALL ECO's", 2070, test_part)

if __name__ == "__main__":
    main()

