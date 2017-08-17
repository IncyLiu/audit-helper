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

resoursepath = r'D:\Dropbox (Linci Work)\MyPrograms\Python 3\audit-helper\details.xls'

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
        sheet.write_merge(1, 1, i, i, '年初')
        sheet.write_merge(1, 1, i+1, i+1, '借方')
        sheet.write_merge(1, 1, i+2, i+2, '贷方')
        sheet.write_merge(1, 1, i+3, i+3, '年末')
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
    print("写入数据成功！请打开test.xls查看结果")
    
testpath = r"C:\Users\Administrator\Desktop\test.xls"
write_excel(ndata, testpath)    