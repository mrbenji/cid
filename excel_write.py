#!/usr/bin/env python3

# standard libraries
import shutil
import pywintypes
import time
import os
import textwrap

# Additional local modules
import cid
from cid_classes import *
import cid_classes

# Third Party Open Source Libs
from xlwings import Workbook, Sheet, Range   # Control Excel via COM. https://pypi.python.org/pypi/xlwings/0.3.4
from colorama import init, Fore, Style   # https://pypi.python.org/pypi/colorama/0.3.3
init()  # for colorama -- initialize functionality

UID_MAP = dict(ast1352="jbs",
               ast1353="jmk",
               ast1382="bdt",
               ast1464="dec",
               ast1560="tmr",
               ast1941="wdm",
               ast2145="lad")


def uid_to_name(uid):
    if uid.lower() in UID_MAP:
        return UID_MAP[uid.lower()]
    else:
        return uid


def calc_first_blank():
    return len(Range('A1').table.value) + 1


def error_text(error):
    hr, msg, exc, arg = error.args
    return textwrap.fill(exc[2])


def save_workbook(workbook, descrip):
    try:
        workbook.save()
    except pywintypes.com_error as error:
        cid.warn_col("\nWARNING: COM error, couldn't write to {}.\n\nCOM error text:".format(descrip))
        print(error_text(error))
        return False

    return True


def open_workbook(workbook_path, descrip, active_tab):

    try:
        wb = Workbook(workbook_path)
    except pywintypes.com_error as error:
        cid.warn_col("\nWARNING: COM error, couldn't open {}.\n\nCOM error text:".format(descrip))
        print(error_text(error))
        return None

    if not save_workbook(wb, descrip):
        return None

    shutil.copy2(workbook_path, workbook_path + ".bak")

    try:
        Sheet(active_tab).activate()
    except pywintypes.com_error:
        time.sleep(1)
        try:
            Sheet(active_tab).activate()
        except pywintypes.com_error as error:
            cid.warn_col("\nWARNING: COM error, couldn't set {} tab {} to active."
                         "\n\nCOM error text:".format(descrip, active_tab))
            print(error_text(error))
            return None

    return wb


def write_config_items_count(eco_path, release_count, close_workbook=True):

    wb = open_workbook(eco_path, "ECO Form", "CoverSheet")

    if not wb:
        return 0

    Range('C16').value = release_count

    if not save_workbook(wb, "ECO Form"):
        return 0

    if close_workbook:
        wb.close()

    return 1


def xlwings_range_to_list_of_parts(range):

    hold_valid_chars = cid_classes.VALID_REV_CHARS
    cid_classes.VALID_REV_CHARS = VALID_AND_INVALID_REV_CHARS

    return_list_of_parts = ListOfParts()

    for row in range.formula:
        return_list_of_parts.add_part(row[0], row[2], row[3])

    cid_classes.VALID_REV_CHARS = hold_valid_chars

    return return_list_of_parts


def write_list_to_pnr(pnrl_path, eco_num, list_of_parts=ListOfParts(), close_workbook=True):

    wb = open_workbook(pnrl_path, "PNR Log", "PN_Rev")

    if not wb:
        return 0

    current_PNR = xlwings_range_to_list_of_parts(Range('A1:D{}'.format(calc_first_blank() - 1)))
    first_row = calc_first_blank()

    current_row = first_row

    user = uid_to_name(os.environ['USERNAME'])
    current_date = time.strftime("%Y-%m-%d")

    not_all_parts_written = False

    # xlwings writes large data blocks much faster if all passed in at once as a 2D list (list of lists)
    add_table = []

    for part_num in list_of_parts.list_of_lists():
        if current_PNR.has_part(part_num[0], part_num[1]):
            cid.warn_col("  Skipping add of {} Rev. {}, was added since last save.".format(part_num[0], part_num[1]))
            not_all_parts_written = True
            continue

        if part_num[2] == eco_num:
            add_table.append([part_num[0], None, part_num[1], part_num[2], None])
            print(Fore.CYAN + "  Adding {} Rev. {}".format(part_num[0], part_num[1]) + Fore.RESET)
        else:
            add_table.append([part_num[0], None, part_num[1], part_num[2],
                              "cid add {} ({} for ECO {})".format(current_date, user, eco_num)])
            print(Fore.CYAN + "  Adding {} Rev. {} (ECO {})".format(part_num[0], part_num[1], part_num[2]) + Fore.RESET)
        current_row += 1

    if not current_row == first_row:
        # pass entire block of data to xlwings for file write
        Range('A{}:E{}'.format(first_row, current_row)).value = add_table
        if not_all_parts_written:
            cid.inf_col('\nINFO: Previous CI additions (listed above as skipped) have now been\n'
                        '      saved to the PN Reserve Log. New additions are not yet saved.')
    else:
        cid.inf_col('\nINFO: Previous CI additions (listed above as skipped) have now been\n'
                    '      saved to the PN Reserve Log.')
        return 0

    if close_workbook:
        if not save_workbook(wb, "PNR Log"):
            return 0
        wb.close()

    return 1

