import helper
import formatter

if __name__ == '__main__':
    input_filepath = '/Users/nrd1012/Projects/momxls/resources/companies_selected.xlsx'
    output_filepath = '/Users/nrd1012/Projects/momxls/resources/done.csv'
    formatter.format_xls(input_filepath, output_filepath, sheet_index=1, name_col=0, address_col=2)

