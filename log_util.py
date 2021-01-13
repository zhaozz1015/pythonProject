#encoding: utf-8
import pandas as pd
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import os
import datetime
from openpyxl import load_workbook

def log_result(type, message):
    result_file_path = 'rest\\result_' + datetime.datetime.now().__format__('%Y-%m-%d') + '.xlsx'
    sheet_name = 'Sheet1'

    if (os.path.exists(result_file_path)):
        df = pd.DataFrame([('', datetime.datetime.now(), type, message)], columns=['', '时间', '类型', '消息'])  # 列表数据转为数据框
        df1 = pd.DataFrame(pd.read_excel(result_file_path, sheet_name=sheet_name))  # 读取原数据文件和表
        writer = pd.ExcelWriter(result_file_path, engine='openpyxl')
        book = load_workbook(result_file_path)
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        df_rows = df1.shape[0]  # 获取原数据的行数
        df.to_excel(writer, sheet_name=sheet_name, startrow=df_rows + 1, index=False, header=False)  # 将数据写入excel中的aa表,从第一个空行开始写
        writer.save()
    else:
        df = pd.DataFrame(columns=['时间', '类型', '消息'])
        df.loc[''] = [datetime.datetime.now(), type, message]
        df.to_excel(result_file_path)
