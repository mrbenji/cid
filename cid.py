import bdt_utils
import argparse
import sys
import io
import openpyxl

VERSION_STRING = "CID v0.6 - 11/18/2014"
HAS_NO_MEDIA = ["scif", "hard_copy", "hardcopy", "synergy"]


def extract_part_nums(filename, all_parts=False):

    if all_parts:
        media_to_skip = []
    else:
        media_to_skip = HAS_NO_MEDIA

    try:
        eco_form = openpyxl.load_workbook(filename)
    except openpyxl.exceptions.InvalidFileException:
        print '\nERROR: Could not open ECO form "{}"\n\n       Is path correct?'.format(filename)
        sys.exit(1)

    pn_sheet = eco_form.get_sheet_by_name('PS1')
    try:
        pn_rows = pn_sheet.rows
    except AttributeError:
        print '\nERROR: ECO form "{}" doesn\'t have a "PS1" tab.'.format(filename)
        sys.exit(1)

    row_num = 0
    current_media = ""
    media_sets = {}
    skip_media = False

    # split PN rows into per-media-type lists
    for row in pn_rows:
        row_num += 1
        if row_num > 4:
            if row_num == 5 and not pn_sheet['G5'].value:
                print "\nERROR @ PS1:G5 - First P/N must have a value in media column."
                exit(1)
            current_media_col = pn_sheet['G'+str(row_num)].value
            prev_media = str(current_media_col).strip().replace(" ", "_").lower()
            if current_media_col and not prev_media == current_media:
                current_media = str(current_media_col).strip().replace(" ", "_").lower()
                if prev_media in media_sets.keys():
                    print "\nERROR @ PS1:G{} - Can't re-use {} after switching to a different type.".format(row_num, current_media)
                    sys.exit(1)

                if not current_media in media_to_skip:
                    media_sets[current_media] = []
            if current_media and not current_media in media_to_skip:
                media_sets[current_media].append(row)

    cid_tables = {}
    current_media = ""

    for set_name in media_sets.keys():
        current_indent_level = 0
        cid_tables[set_name] = []
        pn_table = cid_tables[set_name]

        for row in media_sets[set_name]:
            for cell in row:
                if not pn_sheet['A'+str(cell.row)].value or cell.column not in "ABCFG":
                    continue

                # "Affected Documentation" column
                if cell.column == "A":
                    pn_table.append([])
                    pn_table[-1].append(cell.value)

                # "Cur Rev" column: we skip this if there's a value in "new rev"
                if cell.column == "B":
                    if not cell.value:
                        print "\nERROR on PS1 tab: P/N present in cell A{x}, but B{x} is empty.".format(x=cell.row)
                        exit(1)
                    if not pn_sheet['C'+str(cell.row)].value:
                        pn_table[-1][-1] = pn_table[-1][-1] + " " + "Rev. " + cell.value

                # "New Rev" column
                if cell.column == "C" and cell.value:
                    pn_table[-1][-1] = pn_table[-1][-1] + " " + "Rev. " + cell.value

                # "Description..." column
                if cell.column == "F":
                    if not cell.value:
                        print "\nERROR on PS1 tab: P/N present in cell A{x}, but F{x} is empty.".format(x=cell.row)
                        exit(1)
                    new_indent_level = cell.style.alignment.indent

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
                                skip_media = False
                                hold_row = pn_table.pop()
                                pn_table.append([current_media, ""])
                                pn_table.append(hold_row)

                    # If we're on a row for media we skip, remove entire row from results
                    if skip_media:
                        pn_table.pop()

    return cid_tables
    #return bdt_utils.pretty_table(pn_table, 4)


def print_many_cid_files(contents_id_dump, eol):
    current_media = None

    for line in contents_id_dump.split("\n"):
        if line:
            stripped_line = line.strip()
            if not stripped_line[0].isdigit():
                if current_media:
                    output_file.close()
                current_media = line.strip()
                print "Creating file CONTENTS_ID." + current_media.replace(" ", "_")
                output_file = io.open("CONTENTS_ID." + current_media.replace(" ", "_"), "w", newline=eol)
                continue
        if current_media:
            output_file.write(line+"\n")

    if not output_file.closed:
        output_file.close()


def make_parser():
    """ Construct the command line parser """
    description = VERSION_STRING + " - Create CONTENTS_ID files from PNs on ECO form."
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-v', '--version', action='version', version=VERSION_STRING)
    parser.add_argument("eco_file", type=str, help="eco form filename, w/ full path if not in current dir")

    output_group = parser.add_argument_group('output modes (can be combined)')
    output_group.add_argument('-m', '--print-to-many', action='store_true', default=False,
                        help="print to many files (default)")
    output_group.add_argument('-o', '--print-to-one', action='store_true', default=False,
                        help="print to one file (CONTENTS_ID.all)")
    output_group.add_argument('-s', '--screen-print', action='store_true', default=False,
                        help="print to screen")
    output_group.add_argument('-a', '--all-parts', action='store_true', default=False,
                        help="include PNs that aren't on any media")

    eol_group = parser.add_mutually_exclusive_group()
    eol_group.add_argument('-u', '--unix', action='store_true', default=True,
                       help="use UNIX line endings (default)")
    eol_group.add_argument('-d', '--dos', action='store_true', default=False,
                       help="use DOS line endings")

    return parser


def main():
    parser = make_parser()
    arguments = parser.parse_args(sys.argv[1:])
    # Convert parsed arguments from Namespace to dictionary
    arguments = vars(arguments)

    cid_dumps = extract_part_nums(arguments["eco_file"], arguments["all_parts"])

    if arguments["dos"]:
        eol = '\r\n'
    else: eol = '\n'

    if arguments["screen_print"]:
        for dump in cid_dumps:
            print bdt_utils.pretty_table(cid_dumps[dump], 3)
            print "\n\n"

    if arguments["print_to_one"]:
        print "Creating file CONTENTS_ID.all"
        with io.open("CONTENTS_ID.all", "w", newline=eol) as f:
            for dump in cid_dumps:
                f.write(bdt_utils.pretty_table(cid_dumps[dump], 3))
                f.write(u"\n\n")

    # if no flags were set, default to "print to many" -- if only -o and/or -s were set, don't print to many.
    if not arguments["screen_print"] and not arguments["print_to_one"] and not arguments["print_to_many"]:
        arguments["print_to_many"] = True

    if arguments["print_to_many"]:
        for dump in cid_dumps:
            print_many_cid_files(bdt_utils.pretty_table(cid_dumps[dump], 3), eol)

if __name__ == "__main__":
    main()
