from openpyxl import load_workbook, workbook
from openpyxl import Workbook
import re

def Results_Sum(wb, ws, Title):
    row_title = ['Original GM TC ID', Title]
    ws.append(row_title)
    j = 0
    for sheet in wb:
        print()
        i = 0
        if sheet.title == 'Case need update':
            break
        for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True):
            if not row[0] is None:
                ws.append(row)
                i = i + 1
                j = j + 1
        print(sheet.title + ': ' + str(i))
    print('======================== ' + Title + ' Total:' + str(j) + ' ========================')


#--------------------------------------------- Input Settings -------------------------------------------------
Week = 28
folder = 'C:\\Users\\Yvonne\\Documents\\Test\\'
Order = ['Main', 'Production', 'STR']
#--------------------------------------------------------------------------------------------------------------


output = Workbook()
for Line_Name in Order:
    Title = 'W' + str(Week) + '_' + Line_Name
    wb =  load_workbook(folder + Title +'_Sorted.xlsx')
    ws = output.create_sheet(Line_Name)
    Results_Sum(wb, ws, Title)
output.save(folder + '2021_W' + str(Week) + '.xlsx')