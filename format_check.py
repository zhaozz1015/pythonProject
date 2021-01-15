#encoding: utf-8
import pandas as pd
import log_util
import excel_util

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

    file_url = square[0:square.find('|')]
    check_square = square[square.find('|') + 1 : len(square)]

    print ('file_url:' + file_url + ', check_square:' + check_square)
    excel = pd.read_excel(excel_path, sheet_name=0)
    for i in excel.index.values:
        print (excel.ix[i, 0])

    # for x in params:
    #     print(x)