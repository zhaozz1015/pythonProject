#encoding: utf-8
import log_util

def get_position(mark):
    if (len(mark) <= 1):
        log_util.log_result('错误', '非法的单元格参数' + mark)
        raise Exception('非法的单元格参数' + mark)

    parts = split_alpha_digit(mark)

    return [str(str_to_num(parts[0])), parts[1]]

def split_alpha_digit(mark):
    index = 1
    digit = mark[index:len(mark)]
    while (not digit.isdigit()):
        ++index
        digit = mark[index:len(mark)]

    alpha = mark[0:index]

    return [alpha, digit]

def str_to_num(str):
    strs = {
        'A' : 0,
        'B' : 1,
        'C' : 2,
        'D' : 3,
        'E' : 4,
        'F' : 5,
        'G' : 6,
        'H' : 7,
        'I' : 8,
        'J' : 9,
        'K' : 10,
        'L' : 11,
        'M' : 12,
        'N' : 13,
        'O' : 14,
        'P' : 15,
        'Q' : 16,
        'R' : 17,
        'S' : 18,
        'T' : 19,
        'U' : 20,
        'V' : 21,
        'W' : 22,
        'X' : 23,
        'Y' : 24,
        'Z' : 25,
        'AA' : 26,
        'AB' : 27,
        'AC' : 28,
        'AD' : 29,
        'AE' : 30,
        'AF' : 31,
        'AG' : 32,
        'AH' : 33,
        'AI' : 34,
        'AJ' : 35,
        'AK' : 36,
        'AL' : 37,
        'AM' : 38,
        'AN' : 39,
        'AO' : 40,
        'AP' : 41,
        'AQ' : 42,
        'AR' : 43,
        'AS' : 44,
        'AT' : 45,
        'AU' : 46,
        'AV' : 47,
        'AW' : 48,
        'AX' : 49,
        'AY' : 50,
        'AZ' : 51
    }
    return strs.get(str, None)