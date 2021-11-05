from openpyxl import load_workbook, workbook
from openpyxl import Workbook
import re
import glob

def Results_Sum(wb, ws, path):
    j = 0
    for sheet in wb:
        print()
        i = 0
        if sheet.title == 'Case need update'or sheet.title == 'Summary' or sheet.title == 'Sheet':
            break
        for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):
            if not row[0] is None:
                ws.append(row)
                i = i + 1
                j = j + 1
        print(sheet.title + ': ' + str(i))
    print('====================================================================================='\
     + '\nFrom ' + path\
     + '\nTotal:' + str(j)\
     + '\n=====================================================================================\n')


#--------------------------------------------- Input Settings -------------------------------------------------
Week = 45
folder = 'C:\\Users\\Yvonne\\Documents\\Results'
Order = ['Main', 'STR','SORP']
#--------------------------------------------------------------------------------------------------------------


output = Workbook() # Result Summary
for Line_Name in Order:
    print('# ' + Line_Name + ' #')
    Title = 'W' + str(Week) + '_' + Line_Name
    ws = output.create_sheet(Line_Name) # create sheet for each line
    row_title = ['Original GM TC ID', Title] 
    ws.append(row_title) # add title

    files = glob.glob(folder + '\\Sorted\\' + Title +'_**.xlsx', recursive = True)
    for xls in files :
        wb =  load_workbook(xls)
        Results_Sum(wb, ws, xls)
output.save(folder + '\\All\\2021_W' + str(Week) + '.xlsx')