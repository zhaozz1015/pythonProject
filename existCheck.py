#encoding: utf-8
import pandas as pd
import sys
reload(sys)
sys.setdefaultencoding('utf8')

def exist():
    excel_path = "D:\\Tools\\config.xlsx"
    excel = pd.read_excel(excel_path, sheet_name='existcheck')

    for i in excel.index.values :
        if (excel.ix[i, 0] == '文件') :
            fileExist(excel.ix[i, 1])
        elif (excel.ix[i, 0] == '文件夹') :
            dirExist(excel.ix[i, 1])
        else:
            print('else')

def dirExist(path):
    print('dirExist')

def fileExist(path):
    print('fileExist')