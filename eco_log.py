#!/usr/bin/env python3

# cid.warn_col("\nWARNING: Using local ECO.xlsx for testing, changes will not be live!\n")
# ECO_LOG_PATH = r"c:\cid-tool\cid\ECO.xlsx"

ECO_LOG_PATH = \
    r"\\us.ray.com\SAS\ast\eng\Operations\cm\Internal\Doc_Ctrl\hcm\DOC_CTRL_GENERAL\LOGS\eco_trackinglog\ECO.xlsx"

SRC_PREFIXES = ["040", "123", "219"]
EXE_PREFIXES = ["039", "065", "068", "129", "134", "139", "142", "191", "209", "227"]

# standard libraries
import pywintypes
import textwrap
import sys
import shutil

# Third Party Open Source Libs
from xlwings import Workbook, Sheet, Range  # Control Excel via COM. https://pypi.python.org/pypi/xlwings/0.3.4
import openpyxl                             # https://pypi.python.org/pypi/openpyxl/2.2.0
import six                                  # https://pypi.python.org/pypi/six/1.9.0

import cid
from cid_classes import ListOfParts


def calc_first_blank_xlwings():
    return len(Range('A3').vertical) + 2


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

    sheet_data = Range('A4:V{}'.format(calc_first_blank_xlwings())).value

    for row_num in range(0, len(sheet_data)):
        sheet_data[row_num][0] = round(sheet_data[row_num][0])

    Range('A4:V{}'.format(calc_first_blank_xlwings())).value = sheet_data

    if not save_workbook_xlwings(wb, ECO_LOG_PATH):
        return 0

    return 1


def update_eco_log(wb_path, sheet_name, row_num, part_num_and_rev):

    shutil.copy2(wb_path, wb_path + ".bak")

    wb = open_workbook_xlwings(wb_path)

    if not wb:
        return False

    activate_sheet_xlwings(sheet_name)
    row_to_update = Range('A{r}:E{r}'.format(r=row_num)).value
    row_to_update[3] = "{}".format(part_num_and_rev)
    Range('A{r}:E{r}'.format(r=row_num)).value = row_to_update

    if not save_workbook_xlwings(wb, ECO_LOG_PATH):
        return False

    wb.close()

    return True


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
            # print(eco_num)
            eco_log_data[eco_num] = {}
            eco_log_data[eco_num]["row_num"] = row_num
            eco_log_data[eco_num]["date_assigned"] = row[1].value
            eco_log_data[eco_num]["initiator"] = row[2].value
            eco_log_data[eco_num]["part_number"] = row[3].value
            eco_log_data[eco_num]["project_data"] = row[4].value
            eco_log_data[eco_num]["incorp_date"] = row[5].value

        row_num += 1

    return eco_log_data


def query_one_eco(eco_num):

    eco_log_data = read_openpyxl(ECO_LOG_PATH, "ALL ECO's", "ECO Log")

    if eco_log_data:
        if eco_num in eco_log_data:
            return eco_log_data[eco_num]
        else:
            return None


def check_pns(eco_num, parts=ListOfParts()):

    if isinstance(eco_num, six.string_types):
        eco_num = int(eco_num)
    flat_part_list = parts.flat_list()
    log_data = query_one_eco(eco_num)
    if not log_data:
        return False

    pn_for_log = ""

    # iterate over CI_Sheet rows
    for part_num in flat_part_list:
        if str(log_data["part_number"]).find(part_num) > -1:
            return False

        # The first candidate for the p/n listed on ECO log is the first p/n on the
        # ECO... if it does not have a source or executable prefix, it will only be
        # used if no such p/n's are found on a later row.
        if not pn_for_log:
            pn_for_log = part_num
        else:
            # p/n's with SRC or EXE prefixes are preferred over those without
            if pn_for_log[0:3] not in (SRC_PREFIXES + EXE_PREFIXES) and \
               part_num[0:3] in (SRC_PREFIXES + EXE_PREFIXES):
                pn_for_log = part_num
            # p/n's with EXE prefixes are preferred over those SRC prefixes
            if (pn_for_log[0:3] in SRC_PREFIXES) and (part_num[0:3] in EXE_PREFIXES):
                pn_for_log = part_num

    return_val = update_eco_log(ECO_LOG_PATH, "ALL ECO's", log_data["row_num"], pn_for_log)

    if return_val:
        print("none found.")
        cid.inf_col("Added PN {} to ECO {}'s row in the ECO Log.".format(pn_for_log, eco_num))

    return return_val


def main():
    eco_log_data = read_openpyxl(ECO_LOG_PATH, "ALL ECO's", "ECO Log")

    if eco_log_data:
        parts = ListOfParts()
        parts.add_part("134-128570-01", "J")
        check_pns(11816, parts)
        # eco_log_formulas_to_numbers(ECO_LOG_PATH, "ALL ECO's")
        # update_eco_log(ECO_LOG_PATH, "ALL ECO's", 2070, "123-123456-01", "C")

if __name__ == "__main__":
    main()

