
from xlrd import open_workbook

def is_empty(s):
    return not((not s) or f'{s}'.strip() == '-')


def format_num(s):
    if type(s) == float:
        return str(int(s))
    return s


def get_address(sheet, row_index, t):
    funcs = [_type_1, _type_2, _type_3]
    return funcs[t](sheet, row_index)


def _type_1(sheet, row_index):
    ADDR_NUM = 6
    ADDR_TB = 7
    ADDR_AP = 8
    ADDR_PROVINCE = 9
    ADDR_PC = 10

    addr_num = sheet.cell(row_index, ADDR_NUM).value
    addr_tb = sheet.cell(row_index, ADDR_TB).value
    addr_ap = sheet.cell(row_index, ADDR_AP).value
    addr_province = sheet.cell(row_index, ADDR_PROVINCE).value
    addr_pc = sheet.cell(row_index, ADDR_PC).value

    addr_num = format_num(addr_num) if is_empty(addr_num) else ''
    addr_tb = f' ต.{addr_tb}' if is_empty(addr_tb) else ''
    addr_ap = f' อ.{addr_ap}' if is_empty(addr_ap) else ''
    addr_province = f' จ.{addr_province}' if is_empty(addr_province) else ''
    addr_pc = f' {format_num(addr_pc)}' if is_empty(addr_pc) else ''

    return f'{addr_num}{addr_tb}{addr_ap}{addr_province}{addr_pc}'


def _type_2(sheet, row_index):
    address_col = 1
    address = sheet.cell(row_index, address_col).value
    return address


def _type_3(sheet, row_index):
    ADDR_NUM = 5
    ADDR_MOO = 6
    ADDR_SOI = 7
    ADDR_STREET = 8
    ADDR_TB = 9
    ADDR_AP = 10
    ADDR_PROVINCE = 11
    ADDR_PC = 12

    addr_num = sheet.cell(row_index, ADDR_NUM).value
    addr_moo = sheet.cell(row_index, ADDR_MOO).value
    addr_soi = sheet.cell(row_index, ADDR_SOI).value
    addr_street = sheet.cell(row_index, ADDR_STREET).value
    addr_tb = sheet.cell(row_index, ADDR_TB).value
    addr_ap = sheet.cell(row_index, ADDR_AP).value
    addr_province = sheet.cell(row_index, ADDR_PROVINCE).value
    addr_pc = sheet.cell(row_index, ADDR_PC).value

    addr_num = format_num(addr_num) if is_empty(addr_num) else ''
    addr_moo = f' หมุ่ {format_num(addr_moo)}' if is_empty(addr_moo) else ''
    addr_soi = f' ซอย {addr_soi}' if is_empty(addr_soi) else ''
    addr_street = f' ถนน {addr_street}' if is_empty(addr_street) else ''
    addr_tb = f' ต.{addr_tb}' if is_empty(addr_tb) else ''
    addr_ap = f' อ.{addr_ap}' if is_empty(addr_ap) else ''
    addr_province = f' จ.{addr_province}' if is_empty(addr_province) else ''
    addr_pc = f' {format_num(addr_pc)}' if is_empty(addr_pc) else ''

    return f'{addr_num}{addr_moo}{addr_soi}{addr_street}{addr_tb}{addr_ap}{addr_province}{addr_pc}'