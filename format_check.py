#encoding: utf-8
import pandas as pd
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import log_util
import excel_util
import os

excel_path = "conf\\config.xlsx"

def format_valid():
    excel = pd.read_excel(excel_path, sheet_name='formatcheck')

    for i in excel.index.values :
        if (excel.ix[i, 0] == '文件名称') :
            print('文件名称')
        elif (excel.ix[i, 0] == '单元格值') :
            check_cell_value(str(excel.ix[i, 1]), str(excel.ix[i, 2]), str(excel.ix[i, 3]))
        else:
            log_util.log_result('错误', '不支持的检查类型（规范性）：' + str(excel.ix[i, 0]) + ' | ' + str(excel.ix[i, 1]) + ' | ' + str(excel.ix[i, 2]) + ' | ' + str(
            excel.ix[i, 3]))

def check_cell_value(expression, square, length_limit):
    print('expression:' + expression + ', square:' + square + ', length_limit:' + length_limit)

    params = []  # 配置参数
    if ('<|' in expression and '|>' in expression and expression.find('<|') < expression.find('|>')) :
        excel_exp = expression[expression.find('<|') + 2 : expression.find('|>')]
        sheet_name = excel_exp[0 : excel_exp.find('!')]
        param_val = excel_exp[excel_exp.find('!') + 1 : len(excel_exp)]
        param_star = param_val[0 : param_val.find(':')]
        param_end = param_val[param_val.find(':') + 1 : len(param_val)]

        excel = pd.read_excel(excel_path, sheet_name=sheet_name)

        position_star = excel_util.get_position(param_star)
        position_end = excel_util.get_position(param_end)

        for i in excel.index.values:
            if (i >= int(position_star[1]) - 2 and i <= int(position_end[1]) - 2):
                params.append(excel.ix[i, int(position_star[0])])

    file_url = unicode(square[0:square.find('|')], 'utf-8')
    sheet_index = square[square.find('|') + 1 : square.find('!')]
    check_square = square[square.find('!') + 1 : len(square)]

    param_star = check_square[0: check_square.find(':')]
    param_end = check_square[check_square.find(':') + 1: len(check_square)]

    position_star = excel_util.get_position(param_star)
    position_end = excel_util.get_position(param_end)

    print (str(os.path.isdir(file_url)) + ', file_url:' + file_url + ', sheet_index:' + sheet_index + ', check_square:' + check_square)

    if (os.path.isdir(file_url)):
        files = os.listdir(file_url)
        for file in files:
            check_file_value(expression, params, file_url + file, sheet_index, position_star, position_end)
    else:
        check_file_value(expression, params, file_url, sheet_index, position_star, position_end)

def check_file_value(expression, params, file_url, sheet_index, position_star, position_end):
    check_type = get_check_type(expression)

    print('position_star:' + position_star[0] + ', ' + position_star[1])

    if (check_type == 0) :
        value_add_check(params, file_url, sheet_index, position_star, position_end)
    elif (check_type == 1):
        value_eq_check(params, file_url, sheet_index, position_star, position_end)

def get_check_type(expression):
    check_type = 0 #默认XX+
    if (expression.startswith('<|') and expression.endswith('|>')):
        check_type = 1 #相等

    return check_type

def value_eq_check(params, file_url, sheet_index, position_star, position_end) :
    excel = pd.read_excel(file_url, sheet_name=int(sheet_index))
    valid_value = False

    total_count = 0
    error_count = 0
    for i in excel.index.values:
        if (i >= int(position_star[1]) - 2 and i <= int(position_end[1]) - 2):
            total_count = total_count + 1
            valid_value = False
            value = str(excel.ix[i, int(position_star[0])])
            for x in params:
                if (value == x):
                    valid_value = True

            if (not valid_value):
                error_count = error_count + 1
                log_util.log_result('异常', '[' + value + ']非有效值。' + file_url)

    log_util.log_report('有效值检查，共计' + str(total_count) + '项，其中异常' + str(error_count) + '项。' + file_url)

def value_add_check(params, file_url, sheet_index, position_star, position_end) :
    excel = pd.read_excel(file_url, sheet_name=int(sheet_index))
    valid_value = False

    total_count = 0
    error_count = 0
    for i in excel.index.values:
        if (i >= int(position_star[1]) - 2 and i <= int(position_end[1]) - 2):
            total_count = total_count + 1
            valid_value = False
            value = str(excel.ix[i, int(position_star[0])])
            for x in params:
                if (value.startswith(x)):
                    valid_value = True

            if (not valid_value):
                error_count = error_count + 1
                log_util.log_result('异常', '[' + value + ']非动词+结构。' + file_url)

    log_util.log_report('动词+检查，共计' + str(total_count) + '项，其中异常' + str(error_count) + '项。' + file_url)