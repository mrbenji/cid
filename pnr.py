#!/usr/bin/env python3

import openpyxl  # third party open source library, https://pypi.python.org/pypi/openpyxl/2.2.0
import cid

from cid_classes import *

PNRL_PATH = r"\\us.ray.com\SAS\AST\eng\Operations\CM\Internal\Staff\CM_Submittals\PN_Reserve.xlsm"

# PNRL_PATH = r"\\us.ray.com\SAS\AST\eng\Operations\CM\Internal\Staff\CM_Submittals\PN_Reserve_Copy.xlsm"
# cid.warn_col("\nWARNING: Using PN_Reserve_Copy.xlsm for testing, changes will not be live!\n")

# PNRL_PATH = r"c:\cid-tool\cid\PN_Reserve_copy.xlsm"
# cid.warn_col("\nWARNING: Using local PN_Reserve Log for testing!\n")


def extract_part_nums_pnr():
    """
    Extract part numbers from the part number reserve log, return them as a dict keyed by P/N

    :return: a tuple of values, including...

     - a ListOfParts() for storing the contents of the part number reserve log main worksheet
     - a list of warnings generated during PN Reserve Log extraction, ex. invalid part numbers or revs
     - a ListOfParts() containing duplicate part numbers in the PN Reserve Log
    """

    try:
        # openpyxl is a library for reading/writing Excel files.
        pnr_log = openpyxl.load_workbook(PNRL_PATH)
    except openpyxl.utils.exceptions.InvalidFileException:
        cid.err_col('\nPNR ERROR: Could not open Part Number Reserve Log at path:'
                    '\n       {}'.format(PNRL_PATH))
        cid.exit_app()

    # part number reserve workbook must have a sheet called "PN_Rev"
    pn_sheet = pnr_log.get_sheet_by_name('PN_Rev')
    try:
        pn_rows = pn_sheet.rows
    except AttributeError:
        cid.err_col('\nPNR ERROR: No PN_Rev tab on Part Number Reserve Log at path:'
                    '\n\n     {}'.format(PNRL_PATH))
        cid.exit_app()

    row_num = 0
    pnr_list = ListOfParts()
    pnr_dupe_pn_list = ListOfParts()
    pnr_warnings = []

    if not pn_sheet['A1'].value:
        cid.err_col("\nPNR ERROR: PNR Log does not appear to be valid!"
                    "Cell A1 of {} is blank.".format(PNRL_PATH))
        cid.exit_app()

    for row in pn_rows:
        row_num += 1
        part_num = pn_sheet['A'+str(row_num)].value
        comment_field = pn_sheet['B'+str(row_num)].value
        part_rev = pn_sheet['C'+str(row_num)].value
        eco_num = pn_sheet['D'+str(row_num)].value
        description = pn_sheet['E'+str(row_num)].value

        if part_num and part_rev and eco_num:

            if pnr_list.has_part(part_num, part_rev):
                dupe_pn = "{} Rev. {}".format(part_num, part_rev)
                pnr_warnings.append("PNR WARNING: Duplicate CI {} in PNR Log row {}.".format(dupe_pn, row_num))
                pnr_dupe_pn_list.add_part(part_num, part_rev, eco_num, comment_field)

            try:
                pnr_list.add_part(part_num, part_rev, eco_num, description, comment_field)

            except ValueError:
                if not is_valid_part(part_num) and not (str(comment_field).lower().find("waive") > -1):
                    pnr_warnings.append("PNR WARNING: Skipping PNR Log row {} "
                                        "-- illegal part number.".format(row_num))
                    cid.err_col("\n\nERROR: Illegal part number in PNR Log row {}.\n"
                                "       PN '{}' on ECO {}.\n".format(row_num, part_num, eco_num))
                    input("Press Enter to continue...")

                if not is_valid_rev(part_rev) and not (str(comment_field).lower().find("waive") > -1):
                    pnr_warnings.append("PNR WARNING: Skipping PNR Log row {} "
                                        "-- illegal revision {}.".format(row_num, part_rev))

    return pnr_list, pnr_warnings, pnr_dupe_pn_list
