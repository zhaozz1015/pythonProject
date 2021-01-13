#encoding: utf-8
import pandas as pd
import log_util

def format_valid():
    excel_path = "conf\\config.xlsx"
    excel = pd.read_excel(excel_path, sheet_name='formatcheck')

    for i in excel.index.values :
        print(str(excel.ix[i, 0]) + ' | ' + str(excel.ix[i, 1]) + ' | ' + str(excel.ix[i, 2]) + ' | ' + str(
            excel.ix[i, 3]))

        if (excel.ix[i, 0] == '文件名称') :
            print('文件名称')
        elif (excel.ix[i, 0] == '单元格值') :
            check_cell_value()
        else:
            log_util.log_result('错误', '不支持的检查类型（规范性）：' + str(excel.ix[i, 0]) + ' | ' + str(excel.ix[i, 1]) + ' | ' + str(excel.ix[i, 2]) + ' | ' + str(
            excel.ix[i, 3]))

def check_cell_value():
    print('check_cell_value')