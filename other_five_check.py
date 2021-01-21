#encoding: utf-8
import pandas as pd
import log_util
import os
import excel_util

excel_path = "conf\\config.xlsx"

def check_five():
    excel = pd.read_excel(excel_path, sheet_name='other_five')

    for i in excel.index.values :
        if (excel.ix[i, 0] == 'L5TaskDesc') :
            conditional_not_null_check(str(excel.ix[i, 1]), str(excel.ix[i, 2]), str(excel.ix[i, 3]))
        else:
            log_util.log_result('错误', '不支持的检查类型（五级特别检查）：' + str(excel.ix[i, 0]) + ' | ' + str(excel.ix[i, 1]) + ' | ' + str(excel.ix[i, 2]) + ' | ' + str(
            excel.ix[i, 3]))

def conditional_not_null_check(template_cell_param, cell_value, check_cell_param):
    template_file_url = unicode(template_cell_param[0:template_cell_param.find('|')], 'utf-8')
    template_sheet_index = template_cell_param[template_cell_param.find('|') + 1 : template_cell_param.find('!')]
    template_square = template_cell_param[template_cell_param.find('!') + 1 : len(template_cell_param)]
    check_sheet_index = check_cell_param[check_cell_param.find('|') + 1 : check_cell_param.find('!')]
    check_square = check_cell_param[check_cell_param.find('!') + 1 : len(check_cell_param)]

    # print ('template_file_url:' + template_file_url)
    # print ('template_sheet_index:' + template_sheet_index)
    # print ('template_square:' + template_square)
    # print ('check_sheet_index:' + check_sheet_index)
    # print ('check_square:' + check_square)

    if (os.path.isdir(template_file_url)):
        files = os.listdir(template_file_url)
        for file in files:
            not_null_check(template_file_url + file, template_sheet_index, template_square, cell_value, check_sheet_index, check_square)
    else:
        not_null_check(template_file_url, template_sheet_index, template_square, cell_value, check_sheet_index, check_square)

def not_null_check(template_file_url, template_sheet_index, template_square, cell_value, check_sheet_index, check_square):
    template_excel = pd.read_excel(template_file_url, sheet_name=int(template_sheet_index))


    param_star = template_square[0: template_square.find(':')]
    param_end = template_square[template_square.find(':') + 1: len(template_square)]
    position_star = excel_util.get_position(param_star)
    position_end = excel_util.get_position(param_end)

    squares = check_square.split(',')

    total_count = 0
    error_count = 0
    for i in template_excel.index.values:
        total_count = total_count + 1
        if (template_excel.ix[i, int(position_star[0])] == cell_value):
            check_excel = pd.read_excel(template_file_url, sheet_name=int(check_sheet_index))
            for square in squares:
                check_star = excel_util.get_position(square[0: square.find(':')])
                check_value = str(check_excel.ix[i, int(check_star[0])])
                if (check_value == '' or check_value == 'nan'):
                    error_count = error_count + 1
                    log_util.log_result('异常', '关联非空检查异常【' + excel_util.split_alpha_digit(square[0: square.find(':')])[0] + str(i+2) + '】。' + template_file_url)

    log_util.log_report('关联非空检查，共计' + str(total_count) + '项，其中异常' + str(error_count) + '项。' + template_file_url)