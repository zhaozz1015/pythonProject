#encoding: utf-8
import pandas as pd
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import os

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

        write_excel()

def dirExist(path):
    if (not os.path.exists(path) or not os.path.isdir(path)) :
        print('dir path:' + path)
    else:
        print('dir count:' + str(len(os.listdir(path))))


def fileExist(path):
    if (not os.path.exists(path)) :
        print('file path:' + path)

def write_excel():
    data = {
           'name': ['张三', '李四', '王五', '前六'],
           'age': [11, 12, 13, 14],
           'sex': ['男', '女', '男', '男']
    }
    df = pd.DataFrame(data)
    df.to_excel('D:\\Tools\\new.xlsx')
