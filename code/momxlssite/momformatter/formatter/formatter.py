import momformatter.formatter.helper as helper

MAX_LINES = 4
CHARS_PER_LINE = 35
NUM_COLUMNS = 2
NUM_COLUMNS_IN_BETWEEN = 5
NUM_ROWS_IN_BETWEEN = 2

def format_xls(
        input_filepath,
        output_filepath,
        sheet_index=0,
        name_col=1,
        address_col=3,
        first_row=1,
        chars_per_line=CHARS_PER_LINE,
        max_lines=MAX_LINES):
    raw_data = helper.xls_to_list(input_filepath, sheet_index, name_col, address_col, first_row)

    formatted_data = helper.format_name_and_address(raw_data, chars_per_line, max_lines)
    #formatted_data = helper.format_name_and_address(raw_data[:1], chars_per_line)

    # TODO: make the write_list..'s parameters format_xls's parameters.
    helper.write_list_to_csv(output_filepath, NUM_COLUMNS, NUM_COLUMNS_IN_BETWEEN, NUM_ROWS_IN_BETWEEN, formatted_data)

    print(formatted_data[0])

    for x in formatted_data[0]:
        print("{0} : {1}".format(x, helper.get_real_len(x)))

    #format list
