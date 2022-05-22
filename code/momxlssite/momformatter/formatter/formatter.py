import momformatter.formatter.helper as helper
from functools import reduce

MAX_LINES = 4
CHARS_PER_LINE = 35
NUM_COLUMNS = 2
NUM_COLUMNS_IN_BETWEEN = 5
NUM_ROWS_IN_BETWEEN = 2

electronics_keywords_fragments = [['อิ', 'อี'], ['เลค', 'เลก', 'เล็ค', 'เล็ก'], ['โทร', 'ทรอ'], ['นิค', 'นิก', 'นิคส์', 'นิกส์']]
electronics_keywords = reduce(lambda x, y: [a+b for a in x for b in y], electronics_keywords_fragments)

KEYWORDS = ['โรงงาน', 'ผลิต', 'เกษตร'] + electronics_keywords


def format_xls_to_csv_file(
        input_filepath,
        output_filepath,
        sheet_index=0,
        name_col=1,
        address_col=3,
        description_col=4,
        first_row=1,
        keywords=KEYWORDS,
        chars_per_line=CHARS_PER_LINE,
        max_lines=MAX_LINES):
    raw_data = helper.xls_to_list_with_filter(input_filepath, sheet_index, name_col, address_col, description_col, first_row, keywords)

    formatted_data = helper.format_name_and_address(raw_data, chars_per_line, max_lines)

    # TODO: make the write_list..'s parameters format_xls's parameters.
    helper.write_list_to_csv(output_filepath, NUM_COLUMNS, NUM_COLUMNS_IN_BETWEEN, NUM_ROWS_IN_BETWEEN, formatted_data)

    print("Successfully write to CSV.")
    

def format_xls_to_csv_list(
        input_filepath,
        sheet_index=0,
        name_col=1,
        address_col=3,
        description_col=4,
        first_row=1,
        keywords=KEYWORDS,
        chars_per_line=CHARS_PER_LINE,
        max_lines=MAX_LINES):
    raw_data = helper.xls_to_list_with_filter(input_filepath, sheet_index, name_col, address_col, description_col, first_row, keywords)

    formatted_data = helper.format_name_and_address(raw_data, chars_per_line, max_lines)

    # TODO: make the write_list..'s parameters format_xls's parameters.
    lines = helper.create_writable_list(NUM_COLUMNS, NUM_COLUMNS_IN_BETWEEN, NUM_ROWS_IN_BETWEEN, formatted_data)

    print("Successfully create CSV list.")
    return lines


if __name__ == '__main__':
    input_filepath = '/Users/nrd1012/Projects/momxls/resources/new_raw.xlsx'
    output_filepath = '/Users/nrd1012/Projects/momxls/resources/done.csv'
    format_xls_to_csv_file(input_filepath, output_filepath)

    # print(KEYWORDS)
    # print(len(electronics_keywords))
