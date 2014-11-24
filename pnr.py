import sys
import openpyxl           # third party open source library, https://openpyxl.readthedocs.org/en/latest/
from cid_classes import *

#PNRL_PATH = r"\\us.ray.com\SAS\AST\eng\Operations\CM\Internal\Staff\CM_Submittals\PN_Reserve.xlsm"
PNRL_PATH = "PN_Reserve_copy.xlsm"


def extract_part_nums_pnr():
    """
    Extract part numbers from the part number reserve log, return them as a dict keyed by P/N

    :return: contents of part number reserve log main worksheet, formatted as a dict.  Values are
    lists of dicts {rev:ECO}, keys are base p/n.
    """

    try:
        # openpyxl is a library for reading/writing Excel files.
        pnr_log = openpyxl.load_workbook(PNRL_PATH)
    except openpyxl.exceptions.InvalidFileException:
        print '\nERROR: Could not open Part Number Reserve Log at path:' \
              '\n       {}'.format(PNRL_PATH)
        exit(1)

    # part number reserve workbook must have a sheet called "PN_Rev"
    pn_sheet = pnr_log.get_sheet_by_name('PN_Rev')
    try:
        pn_rows = pn_sheet.rows
    except AttributeError:
        print '\nERROR: No PN_Rev tab on Part Number Reserve Log at path:' \
              '\n\n     {}'.format(PNRL_PATH)
        exit(1)

    row_num = 0
    part_number_dict = {}

    if not pn_sheet['A1'].value:
        print "\nERROR: PNR Log, cell A1 - first cell of part number reserve form is blank."
        exit(1)

    for row in pn_rows:
        row_num += 1
        if pn_sheet['A'+str(row_num)].value and pn_sheet['C'+str(row_num)].value and pn_sheet['D'+str(row_num)].value:
            current_pn = pn_sheet['A'+str(row_num)].value
            if not is_valid_part(current_pn):
                print "\nERROR (PNR LOG): Cell A{} contains an invalid part number.".format(row_num)
                exit(1)
            current_rev = pn_sheet['C'+str(row_num)].value
            if not is_valid_rev(current_rev):
                print "\nERROR (PNR LOG): Cell C{} contains an invalid revision.".format(row_num)
                exit(1)
            current_eco = pn_sheet['D'+str(row_num)].value
            if not current_pn in part_number_dict:
                part_number_dict[current_pn] = {}

            part_number_dict[current_pn][current_rev] = current_eco

    return part_number_dict
