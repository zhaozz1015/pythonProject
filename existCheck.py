#encoding: utf-8
import pandas as pd
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import os
import log_util

def exist():
    excel_path = "conf\\config.xlsx"
    excel = pd.read_excel(excel_path, sheet_name='existcheck')

    for i in excel.index.values :
        if (excel.ix[i, 0] == '文件') :
            fileExist(excel.ix[i, 1])
        elif (excel.ix[i, 0] == '文件夹') :
            dirExist(excel.ix[i, 1])
        else:
            log_util.log_result('错误', '不支持的检查类型（是否存在）：' + excel.ix[i, 0])

def dirExist(path):
    if (not os.path.exists(path) or not os.path.isdir(path)) :
        log_util.log_result('异常', '文件夹不存在：' + path)
    else:
        if (len(os.listdir(path)) == 0):
            log_util.log_result('异常', '文件夹为空：' + path)

def fileExist(path):
    if (not os.path.exists(path)) :
        log_util.log_result('异常', '文件不存在：' + path)