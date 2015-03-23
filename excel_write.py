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

# Third Party Open Source Libs
from xlwings import Workbook, Sheet, Range   # Control Excel via COM. https://pypi.python.org/pypi/xlwings/0.3.4
from colorama import init, Fore, Style   # https://pypi.python.org/pypi/colorama/0.3.3
init()  # for colorama -- initialize functionality

UID_MAP = dict(ast1352="jbs", ast1382="bdt", ast1941="wdm", ast2145="lad", ast1464="dec", ast1560="tmr", ast1353="jmk")


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
    shutil.copy2(workbook_path, workbook_path + ".bak")

    try:
        wb = Workbook(workbook_path)
    except pywintypes.com_error as error:
        cid.warn_col("\nWARNING: COM error, couldn't open {}.\n\nCOM error text:".format(descrip))
        print(error_text(error))
        return None

    if not save_workbook(wb, descrip):
        return None
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


def write_list_to_pnr(pnrl_path, eco_num, list_of_parts=ListOfParts(), close_workbook=True):

    wb = open_workbook(pnrl_path, "PNR Log", "PN_Rev")

    if not wb:
        return 0

    first_row = calc_first_blank()
    current_row = first_row

    user = uid_to_name(os.environ['USERNAME'])
    current_date = time.strftime("%Y-%m-%d")

    # xlwings writes large data blocks much faster if all passed in at once as a 2D list (list of lists)
    add_table = []

    for part_num in list_of_parts.list_of_lists():
        if part_num[2] == eco_num:
            add_table.append([part_num[0], None, part_num[1], part_num[2], None])
            print(Fore.CYAN + "  Adding {} Rev. {}".format(part_num[0], part_num[1]) + Fore.RESET)
        else:
            add_table.append([part_num[0], None, part_num[1], part_num[2],
                              "cid add {} ({} for ECO {})".format(current_date, user, eco_num)])
            print(Fore.CYAN + "  Adding {} Rev. {} (ECO {})".format(part_num[0], part_num[1], part_num[2]) + Fore.RESET)
        current_row += 1

    # pass entire block of data to xlwings for file write
    Range('A{}:E{}'.format(first_row, current_row)).value = add_table

    if not save_workbook(wb, "PNR Log"):
        return 0

    if close_workbook:
        wb.close()

    return 1

