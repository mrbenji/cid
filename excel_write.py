#!/usr/bin/env python3

# standard libraries
import shutil
import pywintypes
import time

# Additional local modules
import cid
from cid_classes import *

# Third Party Open Source Libs
from xlwings import Workbook, Sheet, Range   # Control Excel via COM. https://pypi.python.org/pypi/xlwings/0.3.4
from colorama import init, Fore, Style   # https://pypi.python.org/pypi/colorama/0.3.3
init()  # for colorama -- initialize functionality


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


def write_list_to_pnr(pnrl_path, list_of_parts=ListOfParts(), close_workbook=True):
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

    current_row = calc_first_blank()
    for part_num in list_of_parts.list_of_lists():
        Range('A{r}-D{r}'.format(r=current_row)).value = [part_num[0], None, part_num[1], part_num[2]]
        # Range('C{}'.format(current_row)).value = part_num[1]
        # Range('D{}'.format(current_row)).value = part_num[2]
        print(Fore.CYAN + "  Added {} Rev. {} (ECO {})".format(part_num[0], part_num[1], part_num[2]) + Fore.RESET)
        current_row += 1

    wb.save()
    if close_workbook:
        wb.close()

    return 1

