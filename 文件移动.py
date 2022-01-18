import openpyxl as xl
import os
import shutil

# workBook = xl.load_workbook('D:\\课程文件\\气候变化\\MK值.xlsx')
# sheet = workBook['Sheet2']
# station = []
# for i in range(2,70):
#     station.append(sheet.cell(row=i, column=5).value)

f = open('D:\\课程文件\\气候变化\\冬季插补.txt','r')
station = f.readlines()

file = os.getcwd()
old_path = ['D:\\课程文件\\气候变化\\HDMS2016\\database\\冬季气温\\' + item[5:].strip('\n') + '.txt' for item in station]
for files in old_path:
    try:
        shutil.copy(files, file)
    except:
        print(files)

# file = 'D:\\课程文件\\气候变化\\HDMS2016\\result\\origindata\\春季\\'
# old_path = [file + item[5:].strip('\n') + '春季平均气温原始数据' + '.txt' for item in station]
# for files in old_path:
#     try:
#         os.remove(files)
#     except:
#         print(files)