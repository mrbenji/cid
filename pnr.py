import openpyxl           # third party open source library, https://openpyxl.readthedocs.org/en/latest/
import sys

from cid_classes import *

PNRL_PATH = r"\\us.ray.com\SAS\AST\eng\Operations\CM\Internal\Staff\CM_Submittals\PN_Reserve.xlsm"
#PNRL_PATH = "PN_Reserve_copy.xlsm"


def extract_part_nums_pnr():
    """
    Extract part numbers from the part number reserve log, return them as a dict keyed by P/N

    :return: a tuple of values, including...

     - contents of part number reserve log main worksheet, formatted as a dict.
       Values are lists of dicts {rev:ECO}, keys are base p/n.

     - a list of warnings generated during PN Reserve Log extraction, ex. invalid part numbers or revs
    """

    try:
        # openpyxl is a library for reading/writing Excel files.
        pnr_log = openpyxl.load_workbook(PNRL_PATH)
    except openpyxl.exceptions.InvalidFileException:
        print '\nPNR ERROR: Could not open Part Number Reserve Log at path:' \
              '\n       {}'.format(PNRL_PATH)
        sys.exit(1)

    # part number reserve workbook must have a sheet called "PN_Rev"
    pn_sheet = pnr_log.get_sheet_by_name('PN_Rev')
    try:
        pn_rows = pn_sheet.rows
    except AttributeError:
        print '\nPNR ERROR: No PN_Rev tab on Part Number Reserve Log at path:' \
              '\n\n     {}'.format(PNRL_PATH)
        sys.exit(1)

    row_num = 0
    pnr_list = ListOfParts()
    pnr_warnings = []

    if not pn_sheet['A1'].value:
        print "\nPNR ERROR: PNR Log does not appear to be valid!" \
              "Cell A1 of {} is blank.".format(PNRL_PATH)
        sys.exit(1)

    for row in pn_rows:
        row_num += 1
        part_num = pn_sheet['A'+str(row_num)].value
        part_rev = pn_sheet['C'+str(row_num)].value
        eco_num = pn_sheet['D'+str(row_num)].value

        if part_num and part_rev and eco_num:

            try:
                pnr_list.add_part(part_num, part_rev, eco_num)

            except ValueError:
                if not is_valid_part(part_num):
                    pnr_warnings.append(u"PNR WARNING: Skipping PNR Log row {} -- illegal part number.".format(row_num))
                if not is_valid_rev(part_rev):
                    pnr_warnings.append(u"PNR WARNING: Skipping PNR Log row {} -- illegal revision {}.".format(row_num,
                                                                                                           part_rev))

    return pnr_list, pnr_warnings
