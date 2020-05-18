#快速把多个excel里的内容汇总
import xlrd
import xlsxwriter as xlwt
import os.path
import os

start_row = int(input('从第几行开始?>')
file_address = input('文件夹的地址是?>')
workbook = xlwt.Workbook('{}./汇总.xlsx'.format(file_address))
worksheet = workbook.add_worksheet()
n=0
for i in os.listdir('{}./'.format(file_address)):
    if i.endswith('xlsx'):
        file = xlrd.open_workbook(i)
        info = file.sheet_by_index(0)
        for j in range(start_row-1,info.nrows):
            rows=info.row_values(j)
            worksheet.write_row(n,0,rows)
            n+=1
workbook.close()
                