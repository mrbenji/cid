import bdt_utils
import argparse
import sys
import openpyxl


def put_part_nums_in_table (filename):

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

    for row in pn_rows:
        row_num += 1
        for cell in row:
            if row_num < 5 or not pn_sheet['A'+str(cell.row)].value or cell.column not in "ABCFG":
                continue

            if cell.column == "A":
                pn_table.append([])
                pn_table[-1].append(cell.value)

            if cell.column == "B" and not pn_sheet['C'+str(cell.row)].value:
                pn_table[-1][-1] = pn_table[-1][-1] + " " + "Rev. " + cell.value

            if cell.column == "C" and cell.value:
                pn_table[-1][-1] = pn_table[-1][-1] + " " + "Rev. " + cell.value

            if cell.column == "F" and cell.value:
                pn_table[-1][-1] = "  " * int("{:.0f}".format(cell.style.alignment.indent)) + pn_table[-1][-1]
                pn_table[-1].append(cell.value)

            if cell.column == "G" and cell.value:
                if not current_media == cell.value:
                    current_media = cell.value
                    hold_row = pn_table.pop()
                    pn_table.append(["\n\n" + current_media, ""])
                    pn_table.append(hold_row)

    print bdt_utils.pretty_table(pn_table)


def make_parser():
    """ Construct the command line parser """
    description = "CONTENTS_ID Creator"
    parser = argparse.ArgumentParser(description=description)

    parser.add_argument("eco_file", type=str, help="full path to eco form workbook")

    return parser


def main():
    parser = make_parser()
    arguments = parser.parse_args(sys.argv[1:])
    # Convert parsed arguments from Namespace to dictionary
    arguments = vars(arguments)
    put_part_nums_in_table(arguments["eco_file"])

if __name__ == "__main__":
    main()

