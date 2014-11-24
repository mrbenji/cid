VERSION_STRING = "CID v0.20 - 11/24/2014"

import argparse
import sys
import io
import openpyxl  # third party open source library, https://openpyxl.readthedocs.org/en/latest/
from cid_classes import *  # custom object defs & helper functions for this script
import bdt_utils  # Benji's bag-o'-utility-functions
import pnr

# HAS_NO_MEDIA is a list of "media" tags used for P/Ns that are not put on any official media.  By default
# they are skipped during CONTENTS_ID output. Media tags are converted to lowercase, with spaces converted
# to underscores, before they are checked against this list.
HAS_NO_MEDIA = ["scif", "hard_copy", "hardcopy", "synergy"]


def split_sheet_rows_ps1(pn_sheet, pn_rows, media_to_skip, new_pn_only=False):
    """
    Store rows from PS1 tab of spreadsheet into a dictionary of lists of row objects, keyed by media type.

    :param pn_sheet: openpyxl sheet object
    :param pn_rows: openpxyl rows object
    :param media_to_skip: controls which keywords in the media column mark PN blocks to skip
    :param new_pn_only: if set to True, only new parts will be stored in row lists, and only when they first appear
    :return: dictionary with lists of row objects, keyed by media type
    """

    row_num = 0
    current_media = ""
    media_sets = {}

    # split PN rows into per-media-type lists
    for row in pn_rows:
        row_num += 1

        # skip header section
        if row_num > 4:
            if row_num == 5 and not pn_sheet['G5'].value:
                print "\nERROR: Cell G5 - First P/N must have a value in media column."
                exit(1)
            current_media_col = pn_sheet['G' + str(row_num)].value
            prev_media = str(current_media_col).strip().replace(" ", "_").lower()

            # is this a new media type?
            if current_media_col and not prev_media == current_media:
                current_media = str(current_media_col).strip().replace(" ", "_").lower()
                if prev_media in media_sets.keys():
                    print "\nERROR: Cell G{} - Can't re-use {} after switching to a different type.".format(row_num,
                                                                                                            current_media)
                    sys.exit(1)

                # if this is a media type we want a CONTENTS_ID for, crate an empty list for it in the media_sets dict
                if not current_media in media_to_skip:
                    media_sets[current_media] = []

            # if -n/--new-pn-only is set, we need to verify the part is new before adding this row.
            if new_pn_only and current_media:
                if pn_sheet['C' + str(row_num)].value and not pn_sheet['E' + str(row_num)].value:
                    media_sets[current_media].append(row)
                continue

            # if not on skipped media, add the row to the appropriate list in the media_sets dict
            if current_media and not current_media in media_to_skip:
                media_sets[current_media].append(row)

    return media_sets


def extract_part_nums_PS1(filename, all_parts=False, new_pn_only=False, pnr_list=None, pnr_warnings=[]):
    """
    Open ECO spreadsheet, extract part numbers from the PS1 tab

    :param filename: full path to a properly formatted ECO spreadsheet with completed PS1 tab
    :param all_parts: if True, all PNs will be extracted, including those that wouldn't actually go on media
    :param new_pn_only: if True, only new PNs will be extracted, and only the first time the are listed.
    :return: a dict where each value is a table represented by a list of lists, keyed by media type
    """
    if all_parts:
        media_to_skip = []
    else:
        media_to_skip = HAS_NO_MEDIA

    # pnr_verify will be True if a ListOfParts with the contents of the PN Reserve Log is passed in
    pnr_verify = isinstance(pnr_list, ListOfParts)

    try:
        # openpyxl is a library for reading/writing Excel files.
        eco_form = openpyxl.load_workbook(filename)
    except openpyxl.exceptions.InvalidFileException:
        print '\nERROR: Could not open ECO form at path:' \
              '         {}\n\n       Is path correct?'.format(filename)
        sys.exit(1)

    # ECO form workbook must have a sheet called "PS1"
    pn_sheet = eco_form.get_sheet_by_name('PS1')
    try:
        pn_rows = pn_sheet.rows
    except AttributeError:
        print '\nERROR: No "PS1" tab on ECO form at path:' \
              '         {}'.format(filename)
        sys.exit(1)

    # convert pn_sheet.rows into a dict of row object lists, keyed by media keyword
    media_sets = split_sheet_rows_ps1(pn_sheet, pn_rows, media_to_skip, new_pn_only)

    cid_tables = {}
    current_media = ""
    current_pn = ""
    current_rev = ""
    part_numbers_already_used = {}
    old_part_numbers = {}
    missing_from_pnr_warnings_issued = []
    skip_media = False

    for set_name in media_sets.keys():
        current_indent_level = 0
        cid_tables[set_name] = []
        # pn_table is a reference to cid_tables[set_name], not a copy
        pn_table = cid_tables[set_name]

        for row in media_sets[set_name]:
            for cell in row:

                # we only care about certain columns
                if cell.column not in "ABCEFG":
                    continue

                # Basic line validation: if col A is blank, there should be no values in B or G.
                # Allowing lines like this screws everything up.
                if not pn_sheet['A' + str(cell.row)].value:
                    if pn_sheet['B' + str(cell.row)].value:
                        print "ERROR: P/N present in cell B{x}, but A{x} is empty.".format(x=cell.row)
                        exit(1)
                    elif pn_sheet['G' + str(cell.row)].value:
                        print "ERROR: P/N present in cell G{x}, but A{x} is empty.".format(x=cell.row)
                        exit(1)
                    else:
                        continue

                # "Affected Documentation" column
                if cell.column == "A":
                    pn_table.append([])
                    pn_table[-1].append(cell.value)
                    current_pn = cell.value
                    if not is_valid_part(current_pn):
                        print "ERROR: Cell A{x} contains an improperly-formatted part number.".format(x=cell.row)
                        exit(1)

                # "Cur Rev" column: we skip this if there's a value in "new rev"
                if cell.column == "B":
                    if not cell.value:
                        print "ERROR: P/N present in cell A{x}, but B{x} is empty.".format(x=cell.row)
                        exit(1)
                    if not is_valid_rev(cell.value):
                        print "ERROR: Cell B{x} contains an invalid revision.".format(x=cell.row)
                        exit(1)
                    if not pn_sheet['C' + str(cell.row)].value:
                        pn_table[-1][-1] = pn_table[-1][-1] + " " + "Rev. " + cell.value
                        current_rev = cell.value
                        current_pn_plus_rev = current_pn + " Rev. {}".format(cell.value)

                # "New Rev" column
                if cell.column == "C" and cell.value:
                    if not Rev(pn_sheet['B' + str(cell.row)].value).next_rev.name == cell.value:
                        print "WARNING: Row {x} skips a revision between 'Cur Rev' and 'New Rev.'\n" \
                              "         Is this intentional?".format(x=cell.row)
                    if not is_valid_rev(cell.value):
                        print "ERROR: Cell C{x} contains an invalid revision.".format(x=cell.row)
                        exit(1)
                    pn_table[-1][-1] = pn_table[-1][-1] + " " + "Rev. " + cell.value
                    current_rev = cell.value
                    current_pn_plus_rev = current_pn + " Rev. {}".format(cell.value)

                # "ECO" column -- not useful for CONTENTS_ID, but used for form validation.
                if cell.column == "E":

                    # once a new p/n has been listed, subsequent occurrences must be marked
                    # "dup" in the "ECO" column.  Is this a dup not marked "dup?"
                    if current_pn_plus_rev in part_numbers_already_used.keys() and not cell.value == "dup" \
                            and pn_sheet['C' + str(cell.row)].value:
                        print 'WARNING: Row {} has duplicate P/N {},\n         last used on row {} but ' \
                              'not marked "dup"'.format(cell.row, current_pn_plus_rev,
                                                        part_numbers_already_used[current_pn])

                    # Next: is this a p/n marked "dup" that isn't actually a dup?
                    elif current_pn_plus_rev not in part_numbers_already_used.keys() and cell.value == "dup" \
                            and pn_sheet['C' + str(cell.row)].value:
                        print 'WARNING: Row {} has new P/N {}, incorrectly\n         marked ' \
                              'as "dup"'.format(cell.row, current_pn_plus_rev)

                    # If there's no new rev, there must be an ECO listed in the ECO column
                    elif not pn_sheet['C' + str(cell.row)].value:
                        if not cell.value:
                            print 'ERROR: Cell C{x} has no value, so a value must be added to ' \
                                  'empty cell E{x}!'.format(x=cell.row)
                            exit(1)
                        elif not cell.value == "dup":

                            # if -p/--pnr-verify was set, verify old P/N's vs PN Reserve log
                            if pnr_verify:
                                # Warn if the ECO# listed for a released pn/rev doesn't match what's in the PNR Log
                                if pnr_list.has_part(current_pn, current_rev) \
                                        and pnr_list.parts[current_pn].revs[current_rev].eco != str(cell.value):
                                    print 'ERROR: On ECO row {}, {} is marked as being released on \n       ECO {}. ' \
                                          'This conflicts with the PN Reserve Log, where\n       it is marked as ' \
                                          'released on ECO {}.'.format(cell.row,
                                                                       current_pn,
                                                                       cell.value,
                                                                       pnr_list.parts[current_pn].revs[current_rev].eco)
                                    exit(1)
                                else:
                                    # Report if an old pn/rev combo is not in the PNR Log (report only once per pn/rev)
                                    if current_pn_plus_rev not in missing_from_pnr_warnings_issued:
                                        pnr_warnings.append(u"INFO: ECO row {} - released part {} not "
                                                            "in the PNR Log.".format(cell.row, current_pn_plus_rev))
                                        missing_from_pnr_warnings_issued.append(current_pn_plus_rev)


                            # The following block of validation tests keeps track of the ECO numbers recorded
                            # for previously released part/rev combos.

                            # If this is the first instance of this previously-released part, create a placeholder.
                            if not current_pn in old_part_numbers:
                                old_part_numbers[current_pn_plus_rev] = {}

                            # If the same rev of this previously-released part has been listed earlier...
                            if current_rev in old_part_numbers[current_pn_plus_rev]:

                                # When a previously-released part/rev is listed more than once, all instances
                                # should list the same ECO#
                                if old_part_numbers[current_pn_plus_rev][current_rev] != cell.value:
                                    print "ERROR: On row {}, {} is marked as being released on \n       ECO {}. " \
                                          "This conflicts with row {}, where it is marked as \n       released on " \
                                          "ECO {}.".format(cell.row,
                                                           current_pn_plus_rev,
                                                           cell.value,
                                                           part_numbers_already_used[current_pn_plus_rev],
                                                           old_part_numbers[current_pn_plus_rev][current_rev])
                                    exit(1)

                            # store the ECO# listed for the pn/rev on this row
                            old_part_numbers[current_pn_plus_rev][current_rev] = cell.value

                    # Keep track of part numbers already listed on the ECO
                    part_numbers_already_used[current_pn_plus_rev] = "{}".format(cell.row)

                # "Description..." column
                if cell.column == "F":
                    if not cell.value:
                        print "ERROR: P/N present in cell A{x}, but F{x} is empty.".format(x=cell.row)
                        exit(1)
                    new_indent_level = cell.style.alignment.indent

                    # this will only happen on the first line
                    if not current_indent_level:
                        current_indent_level = new_indent_level

                    # test the current (previous row's) indent level vs. this row's indent level
                    indent_reduced = new_indent_level < current_indent_level
                    current_indent_level = new_indent_level

                    # if the description is indented, we need to add spaces to the part number
                    pn_table[-1][-1] = "  " * int("{:.0f}".format(current_indent_level)) + pn_table[-1][-1]

                    # if the indention level was reduced, add blank line to improve legibility
                    if indent_reduced:
                        pn_table[-1][-1] = "\n" + pn_table[-1][-1]
                    pn_table[-1].append(cell.value)

                # "Media" column
                if cell.column == "G":
                    if cell.value:
                        if not current_media == cell.value:
                            current_media = cell.value
                            if current_media.lower() in media_to_skip:
                                skip_media = True
                            else:
                                # if on a new, non-skipped media type, we pre-pend a line with the media type
                                skip_media = False
                                hold_row = pn_table.pop()
                                pn_table.append([current_media, ""])
                                pn_table.append(hold_row)

                    # If we're on a row for media we skip, remove entire row from results
                    if skip_media:
                        pn_table.pop()

    return cid_tables, pnr_warnings


def write_single_cid_file(contents_id_dump, eol):
    """
    Write the contents of a table to a file

    :param contents_id_dump: a table of part numbers, formatted into a multi-line string by bdt.pretty_table()
    :param eol: the end of line format to use, will be \n for UNIX, \r\n for DOS.
    """

    # break contents_id "dump" into a list of lines
    for line in contents_id_dump.split("\n"):

        # only do the following block on non-blank lines
        if line:
            # create a version of line with leading & trailing spaces removed
            stripped_line = line.strip()

            # does this line not start with a part number?  Then it's a media identifier (CD1, Synergy, etc.)
            # that should be used to name the file, but not be written to the file.
            if not stripped_line[0:3].isdigit():
                current_media = line.strip()
                print "Creating file CONTENTS_ID.{}...".format(current_media.replace(" ", "_"))
                output_file = io.open("CONTENTS_ID." + current_media.replace(" ", "_"), "w", newline=eol)
                continue

        try:
            # write line to file, passes along blank lines, too
            output_file.write(line + "\n")
        except NameError:
            print "\nERROR: write_single_cid_file() was passed a dumpfile without a media ID as the first line."
            exit(1)

    if not output_file.closed:
        output_file.close()


def make_parser():
    """
    Construct a command-line parser for the script, using the build-in argparse library

    :return: an argparse parser object
    """
    description = VERSION_STRING + " - Create CONTENTS_ID files from PNs on ECO form."
    parser = argparse.ArgumentParser(description=description)

    # -v/--version, like -h/--help, ignores other arguments and prints requested info
    parser.add_argument('-v', '--version', action='version', version=VERSION_STRING)

    # this is a required argument unless -v or -h were used
    parser.add_argument("eco_file", type=str, help="eco form filename, w/ full path if not in current dir")

    # parser.add_argument_group creates a named subgroup, for better organization on help screen
    output_group = parser.add_argument_group('output modes (can be combined)')
    output_group.add_argument('-m', '--print-to-many', action='store_true', default=False,
                              help="print to many files (default)")
    output_group.add_argument('-o', '--print-to-one', action='store_true', default=False,
                              help="print to one file (CONTENTS_ID.all)")
    # TURNED THIS ARG OFF FOR NOW, DO WE REALLY NEED IT?
    # output_group.add_argument('-s', '--screen-print', action='store_true', default=False,
    # help="print to screen")
    output_group.add_argument('-a', '--all-parts', action='store_true', default=False,
                              help="include PNs that aren't on any media")
    output_group.add_argument('-e', type=str, choices=["unix", "dos"], default="unix",
                              help="set EOL type for files (default is unix)")

    special_group = parser.add_argument_group('special modes')
    special_meg = special_group.add_mutually_exclusive_group()
    special_meg.add_argument('-n', '--new-pn-only', action='store_true', default=False,
                             help="print only new part numbers, to file NEW_PARTS")
    special_meg.add_argument('-p', '--pnr-verify', action='store_true', default=False,
                             help="verify ECO PNs vs. Part Number Reserve Log")
    special_meg.add_argument('-u', '--update-pnr', action='store_true', default=False,
                             help="update Part Number Reserve Log with ECO PNs (future)")

    return parser


def main():
    """
    Command line execution starts here.
    """

    # "plumbing" for argparse, a standard argument parsing library
    parser = make_parser()
    arguments = parser.parse_args(sys.argv[1:])

    # Convert parsed arguments from Namespace to dictionary
    arguments = vars(arguments)

    new_pn_only = arguments["new_pn_only"]

    # -n automatically prints all parts
    all_parts = arguments["all_parts"] or new_pn_only

    pnr_list = None
    pnr_warnings = []
    if arguments["pnr_verify"]:
        pnr_list, pnr_warnings = pnr.extract_part_nums_pnr()

    # Extract ECO spreadsheet PNs in CONTENTS_ID format (returns a dict of multi-line strings, keyed to media type)
    cid_dumps, pnr_warnings = extract_part_nums_PS1(arguments["eco_file"],
                                                    all_parts, new_pn_only, pnr_list, pnr_warnings)

    # Set file output line endings to requested format.  One (and only one) will always be True.  Default is UNIX.
    if arguments["e"] == "dos":
        eol = '\r\n'
    else:
        eol = '\n'

    if pnr_warnings:
        print 'WARNING: Non-fatal issues found in PN Reserve Log validation phase.\n         See file PNR_WARNINGS ' \
              'for details.'
        with io.open("PNR_WARNINGS", "w", newline=eol) as f:
            for warning in pnr_warnings:
                f.write(warning + "\n")

    # TURNED THIS ARG OFF FOR NOW, DO WE REALLY NEED IT?
    # if arguments["screen_print"]:
    #    for dump in cid_dumps:
    #        print bdt_utils.pretty_table(cid_dumps[dump], 3)
    #        print "\n\n"

    if arguments["new_pn_only"]:
        print "\nCreating file NEW_PARTS...",
        with io.open("NEW_PARTS", "w", newline=eol) as f:
            f.write(u"NOTE: This file lists only the new, unique parts on this ECO. Duplicate and previously-released\n"
                    u"parts are not included.  Nesting is preserved (ex. 065s are indented under the first 139 they\n"
                    u"are affiliated with).  Be aware that you may not be seeing all members of a 139/142/etc., since\n"
                    u"previously-released parts, or parts already displayed under earlier 139s/etc., "
                    u"will be missing.\n\n")
            for dump in cid_dumps:
                if cid_dumps[dump]:
                    f.write(bdt_utils.pretty_table(cid_dumps[dump], 3))
                    f.write(u"\n\n")
    elif arguments["update_pnr"]:
        print "\nThis feature is not yet implemented."
    else:
        # Combine all CONTENTS_IDs into one document.  Can be combined with -m and/or -s.
        if arguments["print_to_one"]:
            print "\nCreating file CONTENTS_ID.all...",
            with io.open("CONTENTS_ID.all", "w", newline=eol) as f:
                for dump in cid_dumps:
                    f.write(bdt_utils.pretty_table(cid_dumps[dump], 3))
                    f.write(u"\n\n")

        # if only -o was set, don't print to many.
        if not arguments["print_to_one"] and not arguments["print_to_many"]:
            arguments["print_to_many"] = True

        if arguments["print_to_many"]:
            print "\n"
            for dump in cid_dumps:
                # write_single_cid_file outputs everything after the media type line to a CONTENTS_ID.<media type> file.
                write_single_cid_file(bdt_utils.pretty_table(cid_dumps[dump], 3), eol)


if __name__ == "__main__":
    main()
