#!/usr/bin/python 
# -*- coding: utf-8 -*-   
import os  

def file_name(file_dir):   
    L=[]   
    i = 1
    for root, dirs, files in os.walk(file_dir):  
        for file in files:  
            if os.path.splitext(file)[1] == '.xls':
                print(i,'.\t',file)
                i += 1
                L.append(os.path.join(root, file))  
    return L

#其中os.path.splitext()函数将路径拆分为文件名+扩展名        
print('您当前目录下有如下excel文件：')
files = file_name(os.getcwd())
snum = input('\n请输入想处理文件的序号：')
resoursepath = files[int(snum)-1]
name = input('\n请给输出文件起个名字：')
testpath = name + '.xls'
print('\n程序将运行半分钟至1分钟，运行完后会自动关闭，输出文件将出现在当前目录下')

YEAR = 1
ACCNUM = 8
ACCNAME = 9
DEB = 10
CRE = 11

#读取excel使用(支持03)  
import xlrd 
#写入excel使用(支持03)  
import xlwt
#读取execel使用(支持07)  
from openpyxl import Workbook  
#写入excel使用(支持07)  
from openpyxl import load_workbook
import numpy as np

workbook=xlrd.open_workbook(resoursepath)  
sheets=workbook.sheet_names();
sheet=workbook.sheet_by_name(sheets[0])  

def pick_col(func, i, c = 1):
    result = func(sheet.col_values(i)) 
    result.remove(sheet.cell_value(0,i))
    result.remove(sheet.cell_value(1,i))
    if c:
        result = func(map(str, result))
    return result

account_set = sorted(list(pick_col(set, ACCNUM)))
year_set = sorted(list(pick_col(set, YEAR, 0)))

nrows = sheet.nrows

ndata = {}
for acc in account_set:
    init = 0.0
    end = 0.0
    tdata = []
    
    for year in year_set:
        ydata = []
        credit = 0.0
        debit = 0.0
        
        #遍历sheet1中所有行row       
        for curr_row in range(nrows):
            row = sheet.row_values(curr_row)
            if(row[YEAR]==year)&(str(row[ACCNUM])==acc):
                credit += row[CRE]
                debit += row[DEB]
                
        end = credit + init - debit
        if np.abs(end) < 1e-3:
            end = 0
        tdata.append(init)
        tdata.append(debit)
        tdata.append(credit)      
        tdata.append(end)
        init = end
    ndata[acc] = tdata
    
check_name = dict(zip(pick_col(list, ACCNUM), pick_col(list, ACCNAME)))

main_acc = sorted(list(set([acc[:4]for acc in account_set])))
ylen = len(year_set)
def write_format(sheet, year_set):
    sheet.write_merge(0, 1, 0, 0, '科目代码')
    sheet.write_merge(0, 1, 1, 1, '科目名称')
    i = 2
    for year in year_set:
        sheet.write_merge(0, 0, i, i+3, '%d年'%year)
        sheet.write_merge(1, 1, i, i,  '年初')
        sheet.write_merge(1, 1, i+1, i+1,  '借方' )
        sheet.write_merge(1, 1, i+2, i+2,  '贷方' )
        sheet.write_merge(1, 1, i+3, i+3,  '年末' )
        i+=4  
def write_excel(data, path):    
    wb=xlwt.Workbook()  
    sheet=wb.add_sheet("sheet1",cell_overwrite_ok=True)
    write_format(sheet, year_set)
    i = 2
    for main in main_acc:
        mdata = [0] * ylen * 4
        mi = i
        i += 1
        for acc in account_set:
            if main in acc:
                sheet.write(i, 0, acc)
                sheet.write(i, 1, check_name[acc])
                for j in range(len(data[acc])):
                    mdata[j] += data[acc][j]
                    sheet.write(i, j+2, round(data[acc][j], 2))
                sheet.write(mi, 0, main)
                for j in range(len(data[acc])):
                    sheet.write(mi, j+2, round(mdata[j]))
                i += 1 
        i += 1
    wb.save(path)  
    print( "\n写入数据成功！" )
    
write_excel(ndata, testpath)    