#!/usr/bin/env python3

# standard libraries
import shutil
import pywintypes
import time
import os

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


def write_config_items_count(eco_path, release_count, close_workbook=True):
    shutil.copy2(eco_path, eco_path + ".bak")
    wb = Workbook(eco_path)
    try:
        wb.save()
    except pywintypes.com_error:
        cid.warn_col("\nWARNING: Another user has the ECO file open, could not write to it.")
        return 0
    try:
        Sheet("CoverSheet").activate()
    except pywintypes.com_error:
        time.sleep(1)
        try:
            Sheet("CoverSheet").activate()
        except pywintypes.com_error:
            cid.warn_col("\nWARNING: Having trouble setting ECO file active tab, could not write to it.")
            return 0

    Range('C16').value = release_count

    wb.save()
    if close_workbook:
        wb.close()

    return 1


def write_list_to_pnr(pnrl_path, eco_num, list_of_parts=ListOfParts(), close_workbook=True):
    shutil.copy2(pnrl_path, pnrl_path + ".bak")
    wb = Workbook(pnrl_path)
    try:
        wb.save()
    except pywintypes.com_error:
        cid.warn_col("\nWARNING: Another user has the PNR Log open, could not write to it.")
        wb.close()
        return 0

    except pywintypes.com_error:
        try:
            time.sleep(1)
            Sheet("PN_Rev").activate()
        except pywintypes.com_error:
            cid.warn_col("\nWARNING: Having trouble setting PNR Log active tab, could not write to it.")
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

    wb.save()
    if close_workbook:
        wb.close()

    return 1

