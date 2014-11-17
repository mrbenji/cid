import bdt_utils
import argparse
import sys
import openpyxl

MEDIA_WE_SKIP = ["scif", "hard copy", "hardcopy", "synergy"]

def extract_part_nums (filename):

    try:
        eco_form = openpyxl.load_workbook(filename)
    except openpyxl.exceptions.InvalidFileException:
        print '\nERROR: Could not open ECO form "{}"'.format(filename)
        sys.exit(1)

    pn_sheet = eco_form.get_sheet_by_name('PS1')
    pn_rows = pn_sheet.rows
    row_num = 0
    pn_table = []
    current_media = ""
    skip_media = False

    for row in pn_rows:
        row_num += 1
        for cell in row:
            if row_num < 5 or not pn_sheet['A'+str(cell.row)].value or cell.column not in "ABCFG":
                continue

            # "Affected Documentation" column
            if cell.column == "A":
                pn_table.append([])
                pn_table[-1].append(cell.value)

            # "Cur Rev" column: we skip this if there's a value in "new rev"
            if cell.column == "B" and not pn_sheet['C'+str(cell.row)].value:
                pn_table[-1][-1] = pn_table[-1][-1] + " " + "Rev. " + cell.value

            # "New Rev" column
            if cell.column == "C" and cell.value:
                pn_table[-1][-1] = pn_table[-1][-1] + " " + "Rev. " + cell.value

            # "Description..." column
            if cell.column == "F" and cell.value:
                # is this row for a 139?
                is_139 = pn_table[-1][-1][0:3] in ["139"]

                # if the description is indented, we need to add spaces to the part number
                pn_table[-1][-1] = "  " * int("{:.0f}".format(cell.style.alignment.indent)) + pn_table[-1][-1]

                # if current PN is a 139, we want to print a blank line before it
                if is_139:
                    pn_table[-1][-1] = "\n" + pn_table[-1][-1]
                pn_table[-1].append(cell.value)

            # "Media" column
            if cell.column == "G":
                if cell.value:
                    if not current_media == cell.value:
                        current_media = cell.value
                        if current_media.lower() in MEDIA_WE_SKIP:
                            skip_media = True
                        else:
                            skip_media = False
                            hold_row = pn_table.pop()
                            pn_table.append(["\n\n" + current_media, ""])
                            pn_table.append(hold_row)

                # If we're on a row for media we skip, remove entire row from results
                if skip_media:
                    pn_table.pop()

    return bdt_utils.pretty_table(pn_table, 4)


def print_many_cid_files(contents_id_dump):
    current_media = None

    for line in contents_id_dump.split("\n"):
        if line:
            stripped_line = line.strip()
            if not stripped_line[0].isdigit():
                if current_media:
                    OUTPUT_FILE.close()
                current_media = line.strip()
                print "Creating file CONTENTS_ID." + current_media
                OUTPUT_FILE = open("CONTENTS_ID." + current_media, "w")
                continue
        if current_media:
            OUTPUT_FILE.write(line+"\n")

    if not OUTPUT_FILE.closed:
        OUTPUT_FILE.close()


def make_parser():
    """ Construct the command line parser """
    description = "Extract PNs from ECO form, create CONTENTS_ID files"
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("eco_file", type=str, help="full path to eco form workbook")
    parser.add_argument('-m', '--print-to-many', action='store_true', default=False, help="print to many files (default)")
    parser.add_argument('-o', '--print-to-one', action='store_true', default=False, help="print to one file (CONTENTS_ID.all)")
    parser.add_argument('-s', '--screen-print', action='store_true', default=False, help="print to screen")
    return parser


def main():
    parser = make_parser()
    arguments = parser.parse_args(sys.argv[1:])
    # Convert parsed arguments from Namespace to dictionary
    arguments = vars(arguments)

    contents_id_dump = extract_part_nums(arguments["eco_file"])

    if arguments["screen_print"]:
        print contents_id_dump

    if arguments["print_to_one"]:
        with open("CONTENTS_ID.all", "w") as f:
            f.write(contents_id_dump)

    # if no flags were set, default to "print to many" -- if only -o and/or -s were set, don't print to many.
    if not arguments["screen_print"] and not arguments["print_to_one"] and not arguments["print_to_many"]:
        arguments["print_to_many"]=True

    if arguments["print_to_many"]:
        print_many_cid_files(contents_id_dump)

if __name__ == "__main__":
    main()
