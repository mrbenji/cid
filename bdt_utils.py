def pretty_money(amount):
    """
    Return integer or float as US-currency-formatted string.
    Ex. 1289 -> "$1,289.00" and 75.98 -> "$75.98"
    :param amount: integer or float to format
    :returns: formatted string
    """
    return "${:,.2f}".format(amount)


def pretty_table(data, padding=2):
    """
    "Pretty print" a table to a multi-line string, with columns as narrow as possible.

    :param data: nested list representing a table (list of rows that are each a list of columns)
    :param padding: minimum space to include between columns
    :returns: a multi-line string containing a formatted table
    """
    error_string = "pretty_table() encountered an improperly-formatted table.\n"
    error_string += "Expected a list of rows, with every row a list of the same number of columns.\n"

    return_string = ""

    # sanity check... is this a list of lists?
    if not isinstance(data, list) and isinstance(data[0], list):
        return error_string

    # make max_col_widths a list of as many 0's as there are columns in the table
    max_col_widths = [0] * len(data[0])

    for row in data:

        # every row should have the same number of columns
        if len(row) != len(data[0]):
            return error_string

        col_num = 0

        # if a column is wider in this row than in previous rows, reset this column's max width
        for col in row:
            if len(col) > max_col_widths[col_num]:
                max_col_widths[col_num] = len(col)
            col_num += 1

    for row in data:
        col_num = 0
        for col in row:
            return_string += col.ljust(max_col_widths[col_num]+padding)
            col_num += 1
        return_string = return_string.rstrip() + "\n"

    return return_string

def ul_string(string_to_ul, ul_char="-"):
    """
    Returns a 1-line string in "underlined" form.
    Does not work properly on strings containing "\n"
    :param string_to_ul: input string
    :param ul_char: character to use for "underlining,"
    defaults to "-"
    :returns: The original string + "\n" + one ul_char
    per char of string_to_ul
    """
    return string_to_ul + "\n" + (ul_char * len(string_to_ul))
