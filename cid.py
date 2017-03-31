#!/usr/bin/env python3

VERSION_STRING = "CID v2.23 02/13/2017"

# standard libraries
import argparse
import sys
import io

# Additional local modules factoring out cid functionality
from cid_classes import *    # Defined classes for Part, Rev, ListOfParts
import cid_classes           # Re-import to allow alternate way to access constants
import bdt_utils             # Benji's bag-o'-utility-functions
import pnr                   # Part Number Reserve Log validation functionality
from excel_write import *    # writing
import eco_log

# third party open source packages
import openpyxl                          # https://pypi.python.org/pypi/openpyxl/2.2.0
from unidecode import unidecode          # https://pypi.python.org/pypi/Unidecode/0.04.17
from colorama import init, Fore, Style   # https://pypi.python.org/pypi/colorama/0.3.3
init()  # for colorama -- initialize functionality

# Update this revision when the ECO form is updated
NEWEST_FORM_REV = Rev('B5')
FORM_REV = None

ECO_PATH = ""
CURRENT_ECO = 0

# HAS_NO_MEDIA is a list of "media" tags used for P/Ns that are not put on any official media.  By default
# they are skipped during CONTENTS_ID output. Media tags are converted to lowercase, with spaces converted
# to underscores, before they are checked against this list.
HAS_NO_MEDIA = ["scif", "hardcopy", "synergy"]

# Keep track of whether errors have been found, so if there are errors we can skip warnings and CID generation
ERRORS_FOUND = False

# Set if -c argument passed, facilitates pausing before exit to keep console from closing
CONSOLE_HOLD = False

# column aliases, used to allow different ECO form revisions to use different columns for the same data
AD_COL = 'A'
CR_COL = 'B'
NR_COL = 'C'
ECO_COL = 'D'
DES_COL = 'E'
MT_COL = 'F'
IN_COL = 'G'
PN_SHEET_COLS = 'ABCDEFG'


def form_rev_switches(cover_sheet):
    """
    :param cover_sheet:  openpyxl Sheet object for the ECO cover sheet
    """
    global ECO_COL
    global DES_COL
    global MT_COL
    global IN_COL
    global PN_SHEET_COLS
    global FORM_REV

    if not cover_sheet['A44'].value:
        FORM_REV = Rev('B1')
    else:
        FORM_REV = Rev(cover_sheet['A44'].value)

    if FORM_REV == Rev('B1') or FORM_REV > Rev('B2'):
        ECO_COL = 'E'
        DES_COL = 'F'
        MT_COL = 'G'
        IN_COL = 'H'
        PN_SHEET_COLS = 'ABCEFGH'

    if FORM_REV < NEWEST_FORM_REV:
        warn_col("\nWARNING: ECO CoverSheet cell A44 does not contain '{}' -- you appear to be\n        "
                 " using old ECO form Rev. {}.  Updating to Rev. {} is"
                 " recommended.".format(NEWEST_FORM_REV.name, FORM_REV.name, NEWEST_FORM_REV.name))
        input("\nPress Enter to continue...")

    if FORM_REV > NEWEST_FORM_REV:
        warn_col("\nWARNING: You are using an newer rev of the ECO form ({}) than the\n"
                 "         CID tool supports. The CID tool needs to be updated.".format(FORM_REV.name))
        input("\nPress Enter to attempt processing anyway, or Ctrl-C to abort.")


def split_sheet_rows_ps1(pn_sheet, cover_sheet, pn_rows, media_to_skip, arguments):
    """
    Store rows from CI_Sheet tab of spreadsheet into a dictionary of lists of row objects, keyed by media type.

    :param pn_sheet: openpyxl sheet object
    :param pn_rows: openpxyl rows object
    :param media_to_skip: controls which keywords in the media column mark PN blocks to skip
    :param arguments: argparse command line arguments, formatted into a hash
    :return: a tuple of values, including...
             - dict with lists of row objects, keyed by media type
             - a list of the keys in the returned dict, to preserve the order they're accessed in
    """

    row_num = 0
    current_media = ""
    current_media_type = ""
    part_number_count = 0
    media_sets = {"skipped": []}
    media_set_order = []
    skip_media_set_appended = False
    new_parts = ListOfParts()

    # split PN rows into per-media-type lists
    for row in pn_rows:
        row_num += 1

        # skip header section
        if row_num > 4:
            if row_num == 5 and not pn_sheet[MT_COL + '5'].value:
                err_col("\nERROR: CI_Sheet cell {}5 - First P/N must have a value in media column.".format(MT_COL))
                exit_app()
            current_media_col = pn_sheet[MT_COL + str(row_num)].value

            # if this row contains a note, we ignore it completely
            if str(current_media_col).strip().lower() in ("note", "notes", "md5sum", "ddf"):
                continue

            if pn_sheet[AD_COL + str(row_num)].value:
                part_number_count += 1

            # is this a new media set?
            if current_media_col:
                current_media_type = str(current_media_col).strip().replace("_", "").replace(" ", "").lower()
                curr_rev = pn_sheet[CR_COL + str(row_num)].value
                new_rev = pn_sheet[NR_COL + str(row_num)].value

                if new_rev:
                    current_media = "{}-{}".format(str(pn_sheet[AD_COL + str(row_num)].value).strip(),
                                                   str(pn_sheet[NR_COL + str(row_num)].value).strip())
                elif curr_rev:
                    current_media = "{}-{}".format(str(pn_sheet[AD_COL + str(row_num)].value.strip()),
                                                   str(pn_sheet[CR_COL + str(row_num)].value))

                # if this is a media type we want a CONTENTS_ID for, crate an empty list for it in the media_sets dict
                if current_media_type not in media_to_skip:
                    media_sets[current_media] = []
                    media_set_order.append(current_media)
                else:
                    if not skip_media_set_appended:
                        media_set_order.append("skipped")
                        skip_media_set_appended = True

            if pn_sheet[NR_COL + str(row_num)].value and not pn_sheet[ECO_COL + str(row_num)].value and current_media:
                try:
                    new_parts.add_part(str(pn_sheet[AD_COL + str(row_num)].value).strip(),
                                       str(pn_sheet[NR_COL + str(row_num)].value).strip(),
                                       CURRENT_ECO,
                                       str(pn_sheet[DES_COL + str(row_num)].value).strip()
                                       )
                except ValueError:
                    err_col("\nERROR: CI_Sheet row {} - "
                            '"{} Rev. {}" is not valid.'.format(row_num,
                                                                str(pn_sheet[AD_COL + str(row_num)].value).strip(),
                                                                str(pn_sheet[NR_COL + str(row_num)].value).strip()
                                                                ))
                    print("       If a rev exception was approved, use the -i argument to override.")
                    exit_app()

                # if -n/--new-pn-only is set, we need to verify the part is new before adding row to media set.
                if arguments["new_pn_only"]:
                    media_sets[current_media].append(row)
                    # return to the top of the for loop
                    continue

            # if not on skipped media, add the row to the appropriate list in the media_sets dict
            if not arguments["new_pn_only"] and current_media and current_media_type not in media_to_skip:
                media_sets[current_media].append(row)
            else:
                media_sets["skipped"].append(row)

    print("\n{} total configuration items. {} CIs "
          "were changed.\n".format(part_number_count, new_parts.count))

    config_items_released = cover_sheet['C16'].value

    if FORM_REV > Rev('B2'):
        if not config_items_released:
            inf_col('INFO: Filled in Cover Sheet cell C16 ("configuration items released").\n'
                    '      Was blank, updated to {}.\n'.format(new_parts.count))
            write_config_items_count(ECO_PATH, new_parts.count, close_workbook=False)
        else:
            if not config_items_released == new_parts.count:
                inf_col('INFO: Updated Cover Sheet cell C16 ("configuration items released").\n'
                        '      Was set to {}, corrected to {}.\n'.format(config_items_released, new_parts.count))
                write_config_items_count(ECO_PATH, new_parts.count, close_workbook=False)

    return media_sets, media_set_order, new_parts


def open_eco():

    try:
        # openpyxl is a library for reading/writing Excel files.
        eco_form = openpyxl.load_workbook(ECO_PATH)

    except openpyxl.utils.exceptions.InvalidFileException:
        err_col('\nERROR: Could not open ECO form at path:\n'
                '       {}\n\n       Is path correct?'.format(ECO_PATH))
        exit_app()

    return eco_form


def extract_ps1_tab_part_nums(arguments, pnr_list=None, pnr_warnings=[], pnr_dupe_pn_list=ListOfParts(), eco_form=None):
    """
    Open ECO spreadsheet, extract part numbers from the PS1 tab

    :param arguments: argparse command line arguments, formatted into a hash
    :param pnr_list: the contents of the part number reserve log, contained in a ListOfParts object
    :return: a tuple of values, including...
             - a dict where each value is a table represented by a list of lists, keyed by media type
             - a list of the keys in the returned dict, to preserve the order they're accessed in
             - a list of warnings generated in the PN_Reserve verification pass
    """

    invalid_revs_ok = arguments["invalid_revs"]

    # -n automatically prints *all* new parts, even if they're not on any media
    if arguments["all_parts"] or arguments["new_pn_only"]:
        media_to_skip = []
    else:
        media_to_skip = HAS_NO_MEDIA

    # pnr_verify will be True if a ListOfParts with the contents of the PN Reserve Log is passed in
    pnr_verify = isinstance(pnr_list, ListOfParts)

    # ECO form workbook must have a sheet named "CoverSheet"
    try:
        cover_sheet = eco_form.get_sheet_by_name('CoverSheet')

    except KeyError:

        # one more attempt to find cover sheet, using original name
        try:
            cover_sheet = eco_form.get_sheet_by_name('NewCoverSheet')

        except KeyError:
            err_col('\nERROR: No "CoverSheet" or "NewCoverSheet" tab in ECO form at path:\n'
                    '       {}'.format(ECO_PATH))
            exit_app()

    # Determine version of form being used, adjust constants, etc. accordingly
    form_rev_switches(cover_sheet)

    # ECO form workbook must have a sheet called "CI_Sheet"
    try:
        pn_sheet = eco_form.get_sheet_by_name('CI_Sheet')
        pn_rows = pn_sheet.rows
    except KeyError:

        # one more attempt to find CI/PN sheet, using original name
        try:
            pn_sheet = eco_form.get_sheet_by_name('PS1')
            pn_rows = pn_sheet.rows

        except KeyError:
            err_col('\nERROR: No "CI_Sheet" or "PS1" tab in ECO form at path:\n'
                    '       {}'.format(ECO_PATH))
            exit_app()

    # convert pn_sheet.rows into a dict of row object lists, keyed by media keyword
    media_sets, media_set_order, new_parts = \
        split_sheet_rows_ps1(pn_sheet, cover_sheet, pn_rows, media_to_skip, arguments)

    global ERRORS_FOUND
    global CURRENT_ECO

    cid_tables = {}
    cid_table_order = []
    CURRENT_ECO = str(cover_sheet['S2'].value)
    current_pn = ""
    current_rev = ""
    prev_rev = ""
    part_numbers_already_used = {}
    old_part_numbers = {}
    missing_from_pnr = ListOfParts()
    next_available_revs = []
    skip_media = False

    for set_name in media_set_order:
        current_indent_level = 0
        cid_tables[set_name] = []
        if not set_name == "skipped":
            cid_table_order.append(set_name)

        # pn_table is a reference to cid_tables[set_name], not a copy,
        # so updates to it will be reflected in the original dict
        pn_table = cid_tables[set_name]

        for row in media_sets[set_name]:

            row_num = row[1].row
            # Basic line validation: if Affected Documentation col is blank, there should be no values in
            # the Cur Rev or Media Type columns.  Allowing lines like this screws everything up.
            if not pn_sheet[AD_COL + str(row_num)].value:
                if pn_sheet[CR_COL + str(row_num)].value:
                    err_col("ERROR: Rev present in CI_Sheet cell {}{}, but no part number in cell {}{}.".format(
                        CR_COL, row_num, AD_COL, row_num))
                    ERRORS_FOUND = True
                elif pn_sheet[MT_COL + str(row_num)].value:
                    err_col("ERROR: Media type present in CI_Sheet cell {}{}, but no PN in cell {}{}.".format(
                        MT_COL, row_num, AD_COL, row_num))
                    ERRORS_FOUND = True
                else:
                    continue

            for cell in row:

                # we only care about certain columns
                if cell.column not in PN_SHEET_COLS:
                    continue

                # "Affected Documentation" column
                if cell.column == AD_COL:
                    pn_table.append([])
                    pn_table[-1].append(str(cell.value).strip())
                    current_pn = str(cell.value).strip()

                    if str(current_pn) == "040-129594-00":
                        pass

                    if not is_valid_part(str(current_pn)):
                        err_col("ERROR: CI_Sheet cell {}{} contains an improperly-formatted part number.".format(
                            AD_COL, cell.row))
                        ERRORS_FOUND = True

                # "Cur Rev" column
                if cell.column == CR_COL:
                    if not cell.value:
                        err_col("ERROR: P/N present in CI_Sheet cell {}{}, but no current rev in {}{}.".format(
                            AD_COL, cell.row, CR_COL, cell.row))
                        ERRORS_FOUND = True
                    if not is_valid_rev(str(cell.value).strip()):
                        if invalid_revs_ok:
                            warn_col("WARNING: CI_Sheet cell {}{} contains invalid "
                                     "revision '{}'.".format(CR_COL, cell.row, str(cell.value).strip()))
                            print("         Script execution continuing because the -i argument was used.")
                        else:
                            err_col("ERROR: CI_Sheet cell {}{} contains invalid "
                                    "revision '{}'.".format(CR_COL, cell.row, str(cell.value).strip()))
                            print("       If an exception was approved, use the -i argument to override.")
                            ERRORS_FOUND = True

                    # if there's not a new revision, this is the revision we're using
                    if not pn_sheet[NR_COL + str(cell.row)].value:
                        current_rev = str(cell.value).strip()
                        current_pn_plus_rev = current_pn + " Rev. {}".format(current_rev)

                        # Replace p/n in last table "cell" with pn+revision
                        pn_table[-1][-1] = current_pn_plus_rev

                        # if there's no new rev, there's no need for the prev_rev var, which is only
                        # used for comparing previous rev to new rev
                        prev_rev = ""

                    else:
                        prev_rev = str(cell.value).strip()

                # "New Rev" column
                if cell.column == NR_COL and cell.value:
                    if not is_valid_rev(str(cell.value).strip()):
                        if invalid_revs_ok:
                            warn_col("WARNING: CI_Sheet cell {}{} contains invalid "
                                     "revision '{}'.".format(NR_COL, cell.row, str(cell.value).strip()))
                            print("         Script execution continuing because the -i argument was used.")
                        else:
                            err_col("ERROR: CI_Sheet cell {}{} contains invalid "
                                    "revision '{}'.".format(NR_COL, cell.row, str(cell.value).strip()))
                            print("       If an exception was approved use the -i argument to override.")
                            ERRORS_FOUND = True

                    if is_valid_rev(pn_sheet[CR_COL + str(cell.row).strip()].value):
                        if not Rev(pn_sheet[CR_COL + str(cell.row).strip()].value).next_rev.name \
                                == str(cell.value).strip() and not ERRORS_FOUND:
                            warn_col("WARNING: CI_Sheet cell {}{} lists new rev '{}'.  Expected '{}', the first\n"
                                     "         valid rev after cur rev "
                                     "{} (cell {}{}).".format(NR_COL, cell.row,
                                                              str(cell.value).strip(),
                                                              Rev(pn_sheet[CR_COL + str(cell.row)].value).next_rev.name,
                                                              pn_sheet[CR_COL + str(cell.row)].value,
                                                              CR_COL,
                                                              cell.row
                                                              )
                                     )

                    if pn_sheet[ECO_COL + str(cell.row)].value and str(
                            pn_sheet[ECO_COL + str(cell.row)].value).isdigit():
                        err_col('ERROR: CI_Sheet row {} -- there cannot be both a new rev in {}{}\n       '
                                'and an ECO number in {}{}.'.format(cell.row, NR_COL, cell.row, ECO_COL,
                                                                    cell.row))
                        ERRORS_FOUND = True

                    current_rev = str(cell.value).strip()

                    current_pn_plus_rev = current_pn + " Rev. {}".format(current_rev)

                    if pnr_verify:

                        if pnr_dupe_pn_list.has_part(current_pn, current_rev):
                            err_col("ERROR: CI_Sheet cell {}{} contains PN {}, which is\n       in the PN Reserve Log "
                                    "more than once.\n".format(AD_COL, cell.row, current_pn_plus_rev))
                            ERRORS_FOUND = True

                        # For new parts, error if pn in PNRL and ECO# listed is not the current ECO.
                        if pnr_list.has_part(current_pn, current_rev):
                            if pnr_list.parts[current_pn].revs[current_rev].eco != CURRENT_ECO:
                                err_col("ERROR: CI_Sheet row {} -- new pn {} is marked in the\n       PN Reserve "
                                        "Log as released on ECO {}, not "
                                        "current ECO {}.\n".format(cell.row,
                                        current_pn_plus_rev, pnr_list.parts[current_pn].revs[current_rev].eco,
                                        CURRENT_ECO))
                                warn_text = "Latest PNR rev for {} is {}, next available is " \
                                            "{}.\n".format(pnr_list.parts[current_pn].number,
                                                           pnr_list.parts[current_pn].max_rev.name,
                                                           pnr_list.parts[current_pn].max_rev.next_rev.name)
                                warn_col("       " + warn_text)
                                next_available_revs.append("Row " + str(cell.row) + ": " + warn_text)
                                ERRORS_FOUND = True

                        else:
                            # Report if a new pn/rev combo is not in the PNR Log (report only once per pn/rev)
                            if not missing_from_pnr.has_part(current_pn, current_rev):
                                pnr_warnings.append("ECO WARNING: row {} - CI {} not in"
                                                    " the PNR Log.".format(cell.row, current_pn_plus_rev))
                                missing_from_pnr.add_part(current_pn, current_rev, CURRENT_ECO)

                            # For new parts, warning if new rev doesn't follow previous rev in PNRL
                            if pnr_list.has_part(current_pn, prev_rev):
                                expected_next_rev = pnr_list.parts[current_pn].revs[prev_rev].next_rev.name
                                if (expected_next_rev != current_rev) and is_valid_rev(current_rev):
                                    error_msg = "WARNING: PN Reserve Log lists the prev rev for {} as '{}'.\n" \
                                                "         Expected new rev '{}' in CI_Sheet " \
                                                "cell {}{}, instead of'{}'.".format(current_pn,
                                                                                    prev_rev,
                                                                                    expected_next_rev,
                                                                                    NR_COL,
                                                                                    cell.row,
                                                                                    current_rev
                                                                                    )

                                    pnr_warnings.append(error_msg)
                                    warn_col(error_msg)

                            # For new parts, warning if new rev is not A or 1 and downrev is not in PNRL
                            elif current_rev not in ('A', '1'):
                                error_msg = "WARNING: CI_Sheet row {} lists new CI {}, whose downrev \n" \
                                            "         is not in the PN Reserve Log. " \
                                            "Is the PN correct?".format(cell.row, current_pn_plus_rev)

                                pnr_warnings.append(error_msg)
                                warn_col(error_msg)

                    # Replace p/n in last table "cell" with pn+revision
                    pn_table[-1][-1] = current_pn_plus_rev

                # "ECO" column -- not useful for CONTENTS_ID, but used for form validation.
                if cell.column == ECO_COL:

                    # once a new p/n has been listed, subsequent occurrences must be marked
                    # "dup" in the "ECO" column.  Is this a dup not marked "dup?"
                    if current_pn_plus_rev in list(part_numbers_already_used.keys()) and not cell.value == "dup" \
                            and pn_sheet[NR_COL + str(cell.row)].value:
                        warn_col('WARNING: CI_Sheet row {} has duplicate P/N {}\n         which is not marked "dup." '
                                 'Last used on row {}.'.format(cell.row, current_pn_plus_rev,
                                                               part_numbers_already_used[current_pn_plus_rev]))

                    # Next: is this a p/n marked "dup" that isn't actually a dup?
                    elif current_pn_plus_rev not in list(part_numbers_already_used.keys()) and cell.value == "dup" \
                            and pn_sheet[NR_COL + str(cell.row)].value:
                        warn_col('WARNING: CI_Sheet row {} has new P/N {}, incorrectly\n         marked '
                                 'as "dup"'.format(cell.row, current_pn_plus_rev))

                    # If there's no new rev, there must be an ECO listed in the ECO column
                    elif not pn_sheet[NR_COL + str(cell.row)].value:
                        if not cell.value:
                            if pnr_verify and pnr_list.has_part(current_pn, current_rev) and \
                               pnr_list.parts[current_pn].revs[current_rev].eco != str(cell.value).strip():
                                err_col('ERROR: CI_Sheet row {} lists {}, which the PNR Log\n'
                                        '       lists as released on ECO {}. Cell {}{} should contain '
                                        "'{}'.\n".format(cell.row, current_pn_plus_rev,
                                                         pnr_list.parts[current_pn].revs[current_rev].eco,
                                                         ECO_COL,
                                                         cell.row,
                                                         pnr_list.parts[current_pn].revs[current_rev].eco
                                                         )
                                        )
                                ERRORS_FOUND = True
                            else:
                                err_col('ERROR: No new rev in CI_Sheet cell {}{}, so an ECO# is required in '
                                        'cell {}{}.\n'.format(NR_COL, cell.row, ECO_COL, cell.row))
                                ERRORS_FOUND = True

                        elif not cell.value == "dup":

                            # if -p/--pnr-verify was set, verify old P/N's vs PN Reserve log
                            if pnr_verify:
                                # Error if the ECO# listed for a released pn/rev doesn't match what's in the PNR Log
                                if pnr_list.has_part(current_pn, current_rev):
                                    if pnr_list.parts[current_pn].revs[current_rev].eco != str(cell.value).strip():
                                        err_col('ERROR: On CI_Sheet row {}, {} is marked as being released on \n       '
                                                'ECO {}. This conflicts with the PN Reserve Log, where\n       '
                                                'it is marked as released '
                                                'on ECO {}.'.format(cell.row, current_pn, str(cell.value).strip(),
                                                                    pnr_list.parts[current_pn].revs[current_rev].eco))
                                        ERRORS_FOUND = True
                                else:
                                    # Report if an old pn/rev combo is not in the PNR Log (report only once per pn/rev)
                                    if not missing_from_pnr.has_part(current_pn, current_rev):
                                        pnr_warnings.append("ECO WARNING: row {} - released CI {} not "
                                                            "in the PNR Log.".format(cell.row, current_pn_plus_rev))
                                        missing_from_pnr.add_part(current_pn, current_rev, cell.value)

                            # The following block of validation tests keeps track of the ECO numbers recorded
                            # for previously released part/rev combos.

                            # If this is the first instance of this previously-released part, create a placeholder.
                            if current_pn not in old_part_numbers:
                                old_part_numbers[current_pn_plus_rev] = {}

                            # If the same rev of this previously-released part has been listed earlier...
                            if current_rev in old_part_numbers[current_pn_plus_rev]:

                                # When a previously-released part/rev is listed more than once, all instances
                                # should list the same ECO#
                                if old_part_numbers[current_pn_plus_rev][current_rev] != str(cell.value).strip():
                                    err_col("ERROR: On CI_Sheet row {}, {} is marked as released on \n       ECO {}. "
                                            "This conflicts with row {}, where it is marked as \n       released on "
                                            "ECO {}.".format(cell.row,
                                                             current_pn_plus_rev,
                                                             str(cell.value).strip(),
                                                             part_numbers_already_used[current_pn_plus_rev],
                                                             old_part_numbers[current_pn_plus_rev][current_rev]))
                                    ERRORS_FOUND = True

                            # store the ECO# listed for the pn/rev on this row
                            old_part_numbers[current_pn_plus_rev][current_rev] = cell.value

                    # Keep track of part numbers already listed on the ECO
                    part_numbers_already_used[current_pn_plus_rev] = "{}".format(cell.row)

                # "Description..." column
                if cell.column == DES_COL:
                    if not cell.value:
                        err_col("ERROR: P/N present in CI_Sheet cell {}{}, but description missing in {}{}.".format(
                            AD_COL, cell.row, DES_COL, cell.row))
                        exit_app()
                    new_indent_level = cell.alignment.indent

                    # this will only happen on the first line
                    if not current_indent_level:
                        current_indent_level = new_indent_level

                    # test the current (previous row's) indent level vs. this row's indent level
                    indent_reduced = new_indent_level < current_indent_level
                    current_indent_level = new_indent_level

                    # if the description is indented, we need to add spaces to the part number
                    pn_table[-1][-1] = "  " * int("{:.0f}".format(current_indent_level)) + pn_table[-1][-1]

                    # if the indention level was reduced, add blank line to improve legibility
                    if indent_reduced and current_indent_level == 0:
                        pn_table[-1][-1] = "\n" + pn_table[-1][-1]

                    pn_table[-1].append(unidecode(cell.value))

                # "Media" column
                if cell.column == MT_COL:
                    if cell.value:
                        current_media = str(cell.value).strip()
                        if current_media.lower() in media_to_skip:
                            skip_media = True
                        else:
                            # if on a new, non-skipped media type, we pre-pend a line with the first part number
                            skip_media = False
                            hold_row = pn_table.pop()
                            pn_table.append([current_media + ":" + set_name, ""])
                            pn_table.append(hold_row)

                    # If we're on a row for media we skip, remove entire row from results
                    if skip_media:
                        pn_table.pop()

                # "ISO Name" column
                if cell.column == IN_COL:
                    if cell.value:
                        iso_name_len = len(str(cell.value).replace('.iso', '').strip())
                        if iso_name_len > 16 and not ERRORS_FOUND:
                            warn_col('WARNING: ISO name in CI_Sheet cell {}{} is {} chars. Is vol name <= '
                                     '16 chars?\n'.format(IN_COL, cell.row, iso_name_len))

    if next_available_revs:
        inf_col("NOTE: Info on next available PNR revs was written to NEXT_AVAILABLE_REV.txt")
        with io.open("NEXT_AVAILABLE_REV.txt", "w", newline="\r\n") as f:
            for line in next_available_revs:
                f.write(unidecode(line))
        f.close()

    return cid_tables, cid_table_order, pnr_warnings, missing_from_pnr, new_parts


def write_single_cid_file(contents_id_table, eol):
    """
    Write the contents of a table to a file

    :param contents_id_table: a table of part numbers, formatted into a multi-line string by bdt.pretty_table()
    :param eol: the end of line format to use, will be \n for UNIX, \r\n for DOS.
    """

    stripped_line = ""

    # break contents_id "table" into a list of lines
    for line in contents_id_table.split("\n"):

        # only do the following block on non-blank lines
        if line:
            line = unidecode(line)

            # create a version of line with leading & trailing spaces removed
            stripped_line = line.strip()

            # does this line not start with a part number?  Then it's a media identifier (CD1, Synergy, etc.)
            # that should be used to name the file, but not be written to the file.
            if not stripped_line[0:3].isdigit():
                current_media = stripped_line[stripped_line.find(":") + 1:]
                inf_col("Creating file CONTENTS_ID.{}...".format(current_media.replace(" ", "_")))
                output_file = io.open("CONTENTS_ID." + current_media.replace(" ", "_"), "w", newline=eol)
                continue

        try:
            # write line to file, passes along blank lines, too
            output_file.write(line + "\n")
        except NameError:
            err_col("\nERROR: write_single_cid_file() was passed an empty or improperly-formatted table."
                    "Table contents:\n{}\nstripped_line:"
                    "\n{}".format(str(contents_id_table), stripped_line))
            exit_app(argparse)

    if not output_file.closed:
        output_file.close()


def warn_col(string_for_warning, end="\n"):
    print(Fore.YELLOW + Style.BRIGHT + string_for_warning + Fore.RESET + Style.RESET_ALL, end=end)


def err_col(string_for_error, end="\n"):
    print(Fore.RED + Style.BRIGHT + string_for_error + Fore.RESET + Style.RESET_ALL, end=end)


def inf_col(string_for_info, end="\n"):
    print(Fore.CYAN + string_for_info + Fore.RESET, end=end)


def make_parser():
    """
    Construct a command-line parser for the script, using the build-in argparse library

    :return: an argparse parser object
    """
    description = VERSION_STRING + " - Create CONTENTS_ID files and validate ECO forms"
    parser = argparse.ArgumentParser(description=description)

    # -v/--version, like -h/--help, ignores other arguments and prints requested info
    parser.add_argument('-v', '-V', '--version', action='version', version=Fore.GREEN + VERSION_STRING + Fore.RESET)

    # this is a required argument unless -v or -h were used
    parser.add_argument("eco_file", type=str, help="eco form filename, w/ full path if not in current dir")

    # parser.add_argument_group creates a named subgroup, for better organization on help screen
    output_group = parser.add_argument_group('output modes (can be combined)')
    output_group.add_argument('-o', '--print-to-one', action='store_true', default=False,
                              help='print all CIs to one file, vs. one per media type')
    output_group.add_argument('-a', '--all-parts', action='store_true', default=False,
                              help="include PNs that aren't on any media")
    output_group.add_argument('-i', '--invalid-revs', action='store_true', default=False,
                              help="allow invalid revisions (issue a warning)")
    output_group.add_argument('-n', '--new-pn-only', action='store_true', default=False,
                              help="print only new part numbers, to file NEW_PARTS")
    output_group.add_argument('-c', '--console-hold', action='store_true', default=False,
                              help="pause before exit, to prevent console closure")
    output_group.add_argument('-e', type=str, choices=["unix", "dos"], default="unix",
                              help="set EOL type for files (default is unix)")

    special_group = parser.add_argument_group('special modes')
    special_meg = special_group.add_mutually_exclusive_group()
    special_meg.add_argument('-np', '--no-pnr-verify', action='store_true', default=False,
                             help="do not verify ECO PNs vs. Part Number Reserve Log")

    return parser


def main():
    """
    Command line execution starts here.
    """

    global ERRORS_FOUND
    global CONSOLE_HOLD
    global ECO_PATH

    # "plumbing" for argparse, a standard argument parsing library
    parser = make_parser()
    arguments = parser.parse_args(sys.argv[1:])

    # needs to be here or the version string will print twice when -v is used
    print(Fore.GREEN + VERSION_STRING + "\n" + Fore.RESET)

    # Convert parsed arguments from Namespace to dictionary
    arguments = vars(arguments)

    # If absolute ECO file path wasn't specified, prepend cwd
    if not os.path.isabs(arguments["eco_file"]):
        ECO_PATH = os.getcwd() + '\\' + arguments["eco_file"]
    else:
        ECO_PATH = arguments["eco_file"]

    # Ignore invalid rev letters (ex. I, O, S, Q, X, Z)
    if arguments["invalid_revs"]:
        cid_classes.VALID_REV_CHARS = VALID_AND_INVALID_REV_CHARS

    pnr_list = None
    pnr_warnings = []
    pnr_dupe_pn_list = ListOfParts()

    # pnr_verify should be the opposite of argument "no_pnr_verify"'s value
    pnr_verify = not arguments["no_pnr_verify"]

    if arguments["new_pn_only"]:
        pnr_verify = False

    CONSOLE_HOLD = arguments["console_hold"]

    if pnr_verify:
        print("Parsing PN Reserve Log...", end=" ")
    elif not arguments["new_pn_only"]:
        warn_col('WARNING: In "-np" mode, skipping validation against PN Reserve Log.')

    # if pnr_verify is still set, parse the PNR log.
    if pnr_verify:
        pnr_list, pnr_warnings, pnr_dupe_pn_list = pnr.extract_part_nums_pnr()

    # Set file output line endings to requested format.  One (and only one) will always be True.  Default is UNIX.
    if arguments["e"] == "dos":
        eol = '\r\n'
    else:
        eol = '\n'

    print("Parsing/validating ECO form...")
    # Extract ECO spreadsheet PNs in CONTENTS_ID format (returns a dict of multi-line strings, keyed to media type)
    cid_tables, cid_table_order, pnr_warnings, missing_from_pnr, new_parts = \
        extract_ps1_tab_part_nums(arguments, pnr_list, pnr_warnings, pnr_dupe_pn_list, eco_form=open_eco())

    if ERRORS_FOUND:
        err_col("\nERRORS FOUND: Skipping CONTENTS_ID creation until resolved.")
        exit_app()

    if new_parts.count:
        print("Checking ECO Log entry for PNs from this ECO... ", end="")
        if not eco_log.check_pns(CURRENT_ECO, new_parts):
            print("OK, ECO Log unchanged.\n")
        else:
            print("")

    if pnr_warnings:

        # Warn if a PN/Rev combo not on this ECO is listed on the PNR as being released on this ECO
        pnr_list_for_this_eco = pnr_list.parts_on_eco(CURRENT_ECO)
        for part in pnr_list_for_this_eco.parts:
            if new_parts.has_part(part):
                for rev in pnr_list_for_this_eco.parts[part].revs:
                    if not new_parts.has_part(part, rev):
                        warn_col("WARNING: PNR Log lists {} Rev. {} as released on this ECO,\n"
                                 "         but that Rev is not listed on this ECO's CI_Sheet.".format(part, rev))
            else:
                warn_col("WARNING: PNR Log lists {} as released on this ECO, but\n"
                         "         that PN is not listed on this ECO's CI_Sheet.".format(part))

        missing_from_pnr_count = missing_from_pnr.count
        if missing_from_pnr_count:
            if missing_from_pnr_count == 1:
                inf_col('INFO: 1 CI on your ECO is missing from the PN Reserve Log.\n')
            else:
                inf_col('INFO: {} CIs on your ECO are missing from the PN Reserve Log.\n'.format(missing_from_pnr_count))
            if input("Automatically add missing CI(s) to PN Reserve Log? [y/N] ") in ['Y', 'y']:
                inf_col('INFO: If this seems to hang, switch to the Excel window displaying the\n'
                        '      PN_Reserve, click "No" in the dialog box, then close the logfile\n'
                        '      without saving. Someone else has the PNR Log open.\n')
                if write_list_to_pnr(pnr.PNRL_PATH, CURRENT_ECO, missing_from_pnr, close_workbook=False):
                    print("")
                    for ci in missing_from_pnr.flat_list():
                        # print(Fore.CYAN + "  Added {}".format(ci) + Fore.RESET)
                        pnr_warnings_copy = pnr_warnings
                        pnr_warnings = []
                        for warning in pnr_warnings_copy:
                            if not warning.find('CI {} not in the PNR Log'.format(ci)) > 0:
                                pnr_warnings.append(warning)

                        pnr_warnings.append("INFO: CID added {} to the PN Reserve Log.".format(ci))
                    warn_col("WARNING: The PNR Log changes are not yet saved, to allow for CMEs\n"
                             "         who would like to verify the added CIs first.  Please\n"
                             "         save & close the PNR Log as soon as practical.\n")
                else:
                    inf_col('\nINFO: Nothing new was written to the PN Reserve Log.\n')
        with io.open("PNR_WARNINGS.txt", "w", newline="\r\n") as f:
            for warning in pnr_warnings:
                f.write(unidecode(warning) + "\r\n")
            f.close()
    else:
        if os.path.isfile("PNR_WARNINGS.txt"):
            os.remove("PNR_WARNINGS.txt")

    if arguments["new_pn_only"]:

        if arguments["print_to_one"]:
            print('NOTE: -n (new parts only) will not generate CONTENTS_ID.X files.\n')

        inf_col("Creating file NEW_PARTS, containing only new, unique parts...", end=' ')

        with io.open("NEW_PARTS", "w", newline=eol) as f:
            f.write("NOTE: This file lists only the new, unique parts on this ECO.\n"
                    "Duplicate and previously-released parts are not included.\n\n")
            f.write(new_parts.text_pretty_list(sort=True))

    else:
        # Combine all CONTENTS_IDs into one document.  Can be combined with -m and/or -s.
        if arguments["print_to_one"]:
            inf_col("Creating file CONTENTS_ID.all...\n", end=' ')
            with io.open("CONTENTS_ID.all", "w", newline=eol) as f:
                for table in cid_table_order:
                    f.write(unidecode(bdt_utils.pretty_table(cid_tables[table], 3)))
                    f.write("\n\n")

        # if only -o was set, don't print to many.
        if not arguments["print_to_one"]:
            for table in cid_table_order:
                # write_single_cid_file outputs everything after the media type line to a CONTENTS_ID.<media type> file.
                write_single_cid_file(bdt_utils.pretty_table(cid_tables[table], 3), eol)

    if FORM_REV < Rev("B3"):
        warn_col('\nWARNING: CID did not verify or update the count of changed CIs on the\n'
                 '         ECO cover sheet, because it detected an outdated ECO form.\n'
                 '         Ensure your cover sheet indicates the correct count of {}.'.format(new_parts.count))

    exit_app(0)


def exit_app(exit_code=1):
    if CONSOLE_HOLD:
        input("\nPress Enter to continue...")

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
