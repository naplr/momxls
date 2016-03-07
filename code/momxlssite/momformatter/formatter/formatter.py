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

    # TODO: make the write_list..'s parameters format_xls's parameters.
    helper.write_list_to_csv(output_filepath, NUM_COLUMNS, NUM_COLUMNS_IN_BETWEEN, NUM_ROWS_IN_BETWEEN, formatted_data)

    print("Successfully write to CSV.")


if __name__ == '__main__':
    input_filepath = '/Users/nrd1012/Projects/momxls/resources/companies_selected.xlsx'
    output_filepath = '/Users/nrd1012/Projects/momxls/resources/done.csv'
    formatter.format_xls(input_filepath, output_filepath, sheet_index=1, name_col=0, address_col=2)
