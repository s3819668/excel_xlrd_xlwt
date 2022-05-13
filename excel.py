import xlrd
import os
from xlutils.copy import copy
from xlutils.styles import Styles
#save wb.save('output.xlsx')   //新建wb2 = Workbook()
#走訪每個Excel檔案
excels=[i for i in os.listdir('./') if i [-4:] =='.xls' and '彙整'not in i]
sheets=[i for i in xlrd.open_workbook(excels[0]).sheet_names()]

des=r'./'+excels[0][0:-9]+'彙整.xls'

sheet_table={}
for sheet in sheets:
    cells= [["" for j in range(23)] for i in range(29)]
    for excel in excels:
        reading_excel = xlrd.open_workbook(excel)
        table=reading_excel.sheet_by_name(sheet)
        for row in range(5, table.nrows-2):
            for col in range(2, 23):
                if type(table.cell(row,col).value)==float or type(table.cell(row,col).value)==int:
                    if cells[row][col]=="":
                        cells[row][col]=table.cell(row,col).value
                    else:
                        cells[row][col]+=table.cell(row,col).value
    sheet_table[sheet]=cells
cp_excel=xlrd.open_workbook(excels[0],formatting_info=True)
writing_excel=copy(cp_excel)

for i in range(len(sheets)):
    ws = writing_excel.get_sheet(i)
    reading_excel = xlrd.open_workbook(excels[0])
    table = reading_excel.sheet_by_name(sheets[i])

    for row in range(5, table.nrows - 2):
        for col in range(2, 23):
            ws.write(row,col,sheet_table[sheets[i]][row][col])

writing_excel.save(des)
input()
