from xlrd import open_workbook
from momformatter.formatter.constants import special_characters

def xls_to_list(filepath, sheet_index, name_col, address_col, first_row):
    sheet = open_workbook(filepath).sheet_by_index(sheet_index)

    li = []
    for row_index in range(first_row, sheet.nrows):

        name = sheet.cell(row_index, name_col).value
        address = sheet.cell(row_index, address_col).value
        r = (name, address)

        li.append(r)

    return li

def get_real_len(s):
    total_len = len(s)
    total_exception = 0
    for c in special_characters:
        total_exception += s.count(c)

    return total_len - total_exception

def format_name_and_address(raw_data, chars_per_line, max_lines):
    def format_name(name):
        return "*บริษัท {0}".format(name)

    def format_address(address, chars_per_line):
        tokens = address.split()

        lines = []
        current_len = 0
        current_line = ""
        for token in tokens:
            token_len = get_real_len(token)
            new_len = current_len + token_len + 1 # plus 1 for a space.
            if (new_len > chars_per_line):
                lines.append(current_line)
                current_line = ""
                current_len = 0

            current_line += "{0} ".format(token)
            current_len += token_len + 1

        # send a warning if the #lines > 2
        lines.append(current_line) # append the last line

        if (len(lines) > max_lines):
            lines.insert(0, "##### TOO MANY LINES #####")

        return lines

    li = []
    for item in raw_data:
        formatted_name = format_name(item[0])
        formatted_address = format_address(item[1], chars_per_line)

        formatted_item = [formatted_name]
        for r in formatted_address:
            formatted_item.append(r)

        li.append(formatted_item)

    return li


def write_list_to_csv(filepath, num_columns, num_columns_in_between, num_rows_in_between, li):

    between = '\t' * num_columns_in_between
    print(between)
    lines = []
    # Iterate through items
    for start_index in range(0, len(li), num_columns):
        # Get list of items that should be on the same line (according to #colums) 
        current_set = li[start_index:start_index + num_columns]
        max_lines = max([len(item) for item in current_set])
        for i in range(max_lines):
            line = ""
            for s in current_set:
                line += (s[i] if len(s) > i else "") + between
            line += "\n"
            lines.append(line)
        for i in range(num_rows_in_between):
            lines.append("\n")

    with open(filepath, 'w', encoding='utf-16') as f:
        f.write('\ufeff')
        f.writelines(lines)

