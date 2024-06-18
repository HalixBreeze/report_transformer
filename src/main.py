import os
import pandas as pd
from pandas.io.excel import ExcelWriter

# 日期時間格式EX: 2021/01/07 06:00:00
start_datetime = '2022/05/05 06:00:00'
end_datetime = '2022/05/12 06:00:00'

# get name of the excel files
read_dir = '../data/origin/'
files = os.listdir(read_dir)

for file_name in files:
    read_path = read_dir + file_name
    data_frame = pd.read_excel(read_path)

    # get data I want from an excel file
    data_frame = data_frame.filter(['部門', '姓名', '刷卡日期', '刷卡時間', '進_出'])

    condition1 = (data_frame['刷卡日期'] + ' ' + data_frame['刷卡時間']) >= start_datetime
    condition2 = (data_frame['刷卡日期'] + ' ' + data_frame['刷卡時間']) < end_datetime

    data_frame = data_frame[condition1 & condition2]

    # write data to an excel file
    write_file = '../data/result/(已修改)' + file_name
    # with pd.ExcelWriter(write_file) as writer:
    #     data_frame.to_excel(excel_writer=writer, sheet_name='Table')
    
    data_frame.to_excel(excel_writer=write_file, sheet_name='Table')

    print(write_file + '已產生')