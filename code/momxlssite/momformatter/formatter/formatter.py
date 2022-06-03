
import momformatter.formatter.helper as helper

MAX_LINES = 4
CHARS_PER_LINE = 50
NUM_COLUMNS = 2
NUM_COLUMNS_IN_BETWEEN = 5
NUM_ROWS_IN_BETWEEN = 2

def format_xls_to_csv_file(
        input_filepath,
        output_filepath,
        format_type,
        chars_per_line=CHARS_PER_LINE,
        max_lines=MAX_LINES):

    raw_data = helper.xls_to_list_with_filter(input_filepath, format_type)
    formatted_data = helper.format_name_and_address(raw_data, chars_per_line, max_lines)

    # TODO: make the write_list..'s parameters format_xls's parameters.
    helper.write_list_to_csv(output_filepath, NUM_COLUMNS, NUM_COLUMNS_IN_BETWEEN, NUM_ROWS_IN_BETWEEN, formatted_data)
    print("Successfully write to CSV.")


def format_xls_to_csv_list(
        input_filepath,
        format_type,
        chars_per_line=CHARS_PER_LINE,
        max_lines=MAX_LINES):

    raw_data = helper.xls_to_list_with_filter(input_filepath, format_type)
    formatted_data = helper.format_name_and_address(raw_data, chars_per_line, max_lines)

    # TODO: make the write_list..'s parameters format_xls's parameters.
    lines = helper.create_writable_list(NUM_COLUMNS, NUM_COLUMNS_IN_BETWEEN, NUM_ROWS_IN_BETWEEN, formatted_data)

    print("Successfully create CSV list.")
    return lines


if __name__ == '__main__':
    input_filepath = '/Users/nrd1012/Projects/momxls/resources/raw-1.xlsx'
    output_filepath = '/Users/nrd1012/Projects/momxls/resources/done.csv'
    format_xls_to_csv_file(input_filepath, output_filepath, 0)

    # print(KEYWORDS)
    # print(len(electronics_keywords))
