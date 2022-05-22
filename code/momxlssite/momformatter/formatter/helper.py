
from xlrd import open_workbook
from momformatter.formatter.constants import special_characters
#from constants import special_characters


def xls_to_list_with_filter(filepath, sheet_index, name_col, description_col, first_row, keywords):
    sheet = open_workbook(filepath).sheet_by_index(sheet_index)

    li = []
    for row_index in range(first_row, sheet.nrows):
        description = sheet.cell(row_index, description_col).value

        if not any(keyword in description for keyword in keywords):
            continue

        name = sheet.cell(row_index, name_col).value
        address = get_address(sheet, row_index)
        r = (name, address, description)

        li.append(r)

    return li

ADDR_NUM = 5
ADDR_MOO = 6
ADDR_SOI = 7
ADDR_STREET = 8
ADDR_TB = 9
ADDR_AP = 10
ADDR_PROVINCE = 11
ADDR_PC = 12
def get_address(sheet, row_index):
    def is_empty(s):
        return not((not s) or f'{s}'.strip() == '-')

    addr_num = sheet.cell(row_index, ADDR_NUM).value
    addr_moo = sheet.cell(row_index, ADDR_MOO).value
    addr_soi = sheet.cell(row_index, ADDR_SOI).value
    addr_street = sheet.cell(row_index, ADDR_STREET).value
    addr_tb = sheet.cell(row_index, ADDR_TB).value
    addr_ap = sheet.cell(row_index, ADDR_AP).value
    addr_province = sheet.cell(row_index, ADDR_PROVINCE).value
    addr_pc = sheet.cell(row_index, ADDR_PC).value

    def format_num(s):
        if type(s) == float:
            return str(int(s))
        return s

    addr_num = format_num(addr_num) if is_empty(addr_num) else ''
    addr_moo = f' หมุ่ {format_num(addr_moo)}' if is_empty(addr_moo) else ''
    addr_soi = f' ซอย {addr_soi}' if is_empty(addr_soi) else ''
    addr_street = f' ถนน {addr_street}' if is_empty(addr_street) else ''
    addr_tb = f' ต.{addr_tb}' if is_empty(addr_tb) else ''
    addr_ap = f' อ.{addr_ap}' if is_empty(addr_ap) else ''
    addr_province = f' จ.{addr_province}' if is_empty(addr_province) else ''
    addr_pc = f' {format_num(addr_pc)}' if is_empty(addr_pc) else ''

    return f'{addr_num}{addr_moo}{addr_soi}{addr_street}{addr_tb}{addr_ap}{addr_province}{addr_pc}'


def format_name_and_address(raw_data, chars_per_line, max_lines):
    def _format_name(name):
        # return "*บริษัท {0}".format(name)
        return "{0}".format(name)

    def _format_address(address, chars_per_line):
        tokens = address.split()

        lines = []
        current_len = 0
        current_line = ""
        for token in tokens:
            token_len = _get_real_len(token)
            new_len = current_len + token_len + 1  # plus 1 for a space.
            if new_len > chars_per_line:
                lines.append(current_line)
                current_line = ""
                current_len = 0

            current_line += "{0} ".format(token)
            current_len += token_len + 1

        # append the last line.
        lines.append(current_line)

        # send a warning if the #lines > max_lines.
        if len(lines) > (max_lines - 1):
            lines.insert(0, "##### TOO MANY LINES #####")

        return lines

    li = []
    for item in raw_data:
        formatted_name = _format_name(item[0])
        formatted_address = _format_address(item[1], chars_per_line)

        formatted_item = [formatted_name]
        for r in formatted_address:
            formatted_item.append(r)

        # (formatted_name_address, description)
        li.append((formatted_item, item[2]))

    return li


def create_writable_list(num_columns, num_columns_in_between, num_rows_in_between, li):
    lines = []
    # Iterate through items
    for start_index in range(0, len(li), num_columns):
        # Get list of items that should be on the same line (according to #colums) 
        current_set = li[start_index:start_index + num_columns]
        max_lines = max([len(item[0]) for item in current_set])
        for i in range(max_lines):
            line = []
            for s in current_set:
                name_address = s[0]
                # Add description on the first line.
                line.append((name_address[i] if len(name_address) > i else ''))
                line.append(s[1] if i == 0 else '')
                line.append('')
            lines.append(line)
        for i in range(num_rows_in_between):
            lines.append([])

    return lines


def create_writable_lines(num_columns, num_columns_in_between, num_rows_in_between, li):
    between = '\t' * num_columns_in_between
    print(between)
    lines = []
    # Iterate through items
    for start_index in range(0, len(li), num_columns):
        # Get list of items that should be on the same line (according to #colums) 
        current_set = li[start_index:start_index + num_columns]
        max_lines = max([len(item[0]) for item in current_set])
        for i in range(max_lines):
            line = ""
            for s in current_set:
                name_address = s[0]
                # Add description on the first line.
                line += (name_address[i] if len(name_address) > i else "") + "\t" + (s[1] if i == 0 else "") + between
            line += "\n"
            lines.append(line)
        for i in range(num_rows_in_between):
            lines.append("\n")

    return lines


def write_list_to_csv(filepath, num_columns, num_columns_in_between, num_rows_in_between, li):
    # TODO: switch to use csv package, so we can reuse the create_writeable_list and get rid of create_writable_lines.
    lines = create_writable_lines(num_columns, num_columns_in_between, num_rows_in_between, li)

    with open(filepath, 'w', encoding='utf-16') as f:
        f.write('\ufeff')
        f.writelines(lines)


def _get_real_len(s):
    total_len = len(s)
    total_exception = 0
    for c in special_characters:
        total_exception += s.count(c)

    return total_len - total_exception
