import os
os.system('CLS') 

# from openpyxl import load_workbook
# '''Создаем класс книги'''
# wb = load_workbook("one.xlsx")
# """Подключаемся к активному листу"""
# # ws = wb.active
# res = cell_range = ws['A1':'C10']
# ws1 = wb.create_sheet("Mysheet")
# ws1.title = "New Title"
# wb.save('one.xlsx')
# print("res = ", res[0])
# sheet = wb['Лист1']
# val = sheet['A1'].value
# print("val = ", val)
# vals = [v[0].value for v in sheet.range('A1:A3')]
# print("vals = ", vals)

# import openpyxl

# from openpyxl import Workbook
# wb = Workbook()


# # ws = wb.active

# min_column = wb.active.min_column
# max_column = wb.active.max_column
# min_row = wb.active.min_row
# max_row = wb.active.max_row


# print(f"max_row = {max_row}")
# print(f"max_column = {max_column}")


# import win32com.client

# """Создаем COM объект"""
# Excel = win32com.client.Dispatch("Excel.Application")
# Excel.Visible=1

# """Открываем определенный файл"""
# # wb = Excel.Workbooks.Open(r"C:\Users\vvkhomutskiy\Documents\vxv\code\ExPro\one.xlsx")
# # wb = Excel.Workbooks.Open(r"Кодировка_тест.xlsx")

# """Получаем доступ к активной книге"""
# wb = Excel.ActiveWorkbook

# """Получаем доступ к активному листу"""
# sheet = wb.ActiveSheet

# # iLastRow = Excel.ActiveCell.Row
# # print("iLastRow = ", iLastRow)

# """Получаем значение ячеек построчно"""
# TabList = []
# # for row in range(1, 10+1):
# for row in range(1, 100):
#     RowList = []
#     for col in range(1, 100):
#         if row != "" and col != "":
#             val = sheet.Cells(row, col).value
#             RowList.append(val)
#     TabList.append(RowList)

# # print("TabList = ", TabList)



# """получаем значения цепочки A1:A2"""
# vals = [i[0].value for i in sheet.Range("A1:C3")]

# # print("vals = ", vals)



"""Закроем файл"""
# wb.Close()

"""Закроем COM объект"""
# Excel.Quit()






# import win32com.client

# xlApp = win32com.client.Dispatch("Excel.Application")
# xlApp.Visible=1

# workBook = xlApp.Workbooks.Activ
# print str(workBook.ActiveSheet.Cells(i,1))
# workBook.ActiveSheet.Cells(1, 1).Value = "hello"                
# workBook.Close(SaveChanges=0) 
# xlApp.Quit()



'''Создаем границы ячеек'''
# # sheet3.Cells(16, 2).Borders.ColorIndex = 0
# sheet4.Cells(15, 2).Borders(1).Weight = 3       # Левая граница
# sheet4.Cells(15, 2).Borders(2).Weight = 3       # Правая граница
# sheet4.Cells(15, 2).Borders(3).Weight = 3       # Верхняя граница
# sheet4.Cells(15, 2).Borders(4).Weight = 3       # Нижняя граница
# sheet4.Cells(15, 2).Borders(5).Weight = 3       # Диагональ граница
# sheet4.Cells(15, 2).Borders(6).Weight = 3       # Диагональ граница
# sheet4.Cells(15, 2).Borders.Weight = 3          # Все границы
# sheet4.Range("B5:G7").Borders(7).Weight = 3       # Левая граница крайних ячеек в диапозоне
# sheet4.Range("B5:G7").Borders(8).Weight = 3       # Верхняя граница крайних ячеек в диапозоне
# sheet4.Range("B5:G7").Borders(9).Weight = 3       # Нижняя граница крайних ячеек в диапозоне
# sheet4.Range("B5:G7").Borders(10).Weight = 3      # Правая граница крайних ячеек в диапозоне
# sheet4.Range("B5:G7").Borders(11).Weight = 3      # Вертикальные внутренние границы ячеек в диапозоне
# sheet4.Range("B5:G7").Borders(12).Weight = 3      # Горизонтальные внутренние границы ячеек в диапозоне


# sheet.Range(f"C{StartNomerRow + row}").WrapText = True
# sheet.Range(f"C{StartNomerRow + row}").HorizontalAlignment = 1

# sheet3.Range("C").Orientation = 0
# sheet3.Range("C").AddIndent = False
# sheet3.Range("C").IndentLevel = 0
# sheet3.Range("C").ShrinkToFit = False
# # sheet3.Range("C").ReadingOrder = xlContext
# sheet3.Range("C").MergeCells = False

'''Удаляем строки со сдвигом вверх'''
# sheet3.Rows("7:57").Delete(1)
'''Очистить содержимое строчек'''
# sheet3.Rows("7:57").ClearContents()

# sheetOpti = wb.Worksheets("Options")
# cel = sheetOpti.Range(sheetOpti.Cells(2, 1), sheetOpti.Cells(13, 47))
# cel.Copy()
# sheet.Rows(RowEnd + 1).Activate()
# sheet.Paste()

# '''от правого края влево до правой крайней заполненной ячейки'''
# max_col = sheet.Cells(row, sheet.Columns.Count).End(1).Column

# cellll = sheet.Cells(row, max_col).Interior
# cellll.Pattern = 2
# cellll.PatternColorIndex = 3

# '''от 1ой колонки вправо до первой заполненной ячейки'''
# min_col = sheet.Cells(row+1, 1).End(2).Column
# if min_col == max_col: min_col = 1

# cellll = sheet.Cells(row+1, min_col).Interior
# cellll.Pattern = 2
# cellll.PatternColorIndex = 4

# '''Количество занимаемых таблицей строк'''
# countTab_row = sheet.UsedRange.Rows.Count
# '''Количество занимаемых таблицей колонок'''
# countTab_col = sheet.UsedRange.Columns.Count


# def codeTalie(sheet, StartNomerRow):
#     '''hhh = ''
#     kLuch = []
#     for i in range(len(row1)):
#         if isinstance(row1[i], str):
#             ggg = row1[i]
#             hhh = ggg
#         else:
#             ggg = hhh
#         # ggg = str(row1[i - 1])
#         for xx in [row2[i], row3[i], row4[i], row5[i]]:
#             if isinstance(xx, str):
#                 ggg = ggg + " - " + xx
#         kLuch.append(ggg)
#     # print(f"kLuch = {kLuch}")

#     RewXXX = importda(sheet1, 15, StartNomerColl, 15, EndNomerColl)
#     RewXXXDict = {}
#     for i in range(len(row1)):
#         RewXXXDict[kLuch[i]] = RewXXX[i]
#     fff = RewXXXDict['Глубина отбора пробы. м']
#     print(f"fff = {fff}")'''


'''def rashod(sheet, StartNomerRow, EndNomerRow, xxx, yyy):
#     Сравниваем колонку с класиыикацией от лабораторией и нашу
#     for i in range(StartNomerRow, EndNomerRow + 1):
#         aaa = sheet.Cells(i, xxx).value
#         bbb = sheet.Cells(i, yyy).value
#         if bbb not in aaa:
#             sheet.Cells(i, yyy).Font.Color = -16776961'''


# IGEnd = sheet.Range(sheet.Cells(StartNomerRow + Row, StartNomerColl), sheet.Cells(StartNomerRow + Row, EndNomerColl))
# IGEnd.VerticalAlignment = 2
'''Центрует по горизонтали центр'''
# IGEnd.HorizontalAlignment = 3
'''Центрует по горизонтали влево'''
# IGEnd.HorizontalAlignment = 1


# def Formula_schet(sheet, rowN, colN):
#     # sheet.Cells(rowN, colN).Formula = f"=СЧЁТ(C{rowN - lenX}:C{rowN - 1})"
#     cel = sheet.Cells(rowN, colN)
#     if cel.value != None:
#         # cel.Formula = f"=СЧЁТ(R[{-lenX}]C:R[{-1}]C)"
#         cel.Formula = f"=COUNT(R[{-lenX}]C:R[-1]C)"
#         cel.Copy()
#         sheet.Cells(rowN, colN + 1).Activate()
#         sheet.Paste()

# sheet.Cells(rowXn, 6).Formula = f"=D{rowXn}-E{rowXn}"
# sheet.Cells(rowXn, 7).Formula = f"=(C{rowXn}-E{rowXn})/F{rowXn}"
# sheet.Cells(rowXn, 10).Formula = f"=H{rowXn}/(1+C{rowXn})"
# sheet.Cells(rowXn, 11).Formula = f"=((I{rowXn}-J{rowXn})/I{rowXn})*100"
# sheet.Cells(rowXn, 12).Formula = f"=(I{rowXn}-J{rowXn})/J{rowXn}"
# sheet.Cells(rowXn, 13).Formula = f"=((C{rowXn}*I{rowXn})/L{rowXn})"
# sheet.Cells(rowXn, 32).Formula = f"=AD{rowXn}*AE{rowXn}"
# sheet.Cells(rowXn, 35).Formula = f"=DEGREES(ATAN(RC[-1]))"



'''Формулы (при вставке значений по вертикали использовать кортежи с пустым 2ым значением)
# пример: ("=формула", )"'''
# formula1 = "=RC[-2]-RC[-1]"
# formula2 = "=(RC[-4]-RC[-2])/RC[-1]"
# formula12 = [(formula1, formula2)]*countRow
# cell = sheet.Range(sheet.Cells(StartNomerRow, 6), sheet.Cells(EndNomerRow, 7))
# cell.ClearContents()
# cell.value = formula12
# cell.Font.Color = -1179134

# def ifErr(formula):
#     iferror = f"IFERROR({formula},\"\")"
#     text = f"=IF({iferror}=0,\"\",{iferror})"
#     # print(text)
#     return text

'''если ошибка'''
''' "=IFERROR({formula},\"\")" '''


# rowN = StartNomerRow + Row
# for colN in range(3, EndNomerColl + 1):
#     Formula_schet(sheet, rowN, colN)

# Formula_schet(sheet, rowN, 3)



# import win32com.client

# """Создаем COM объект"""
# Excel = win32com.client.Dispatch("Excel.Application")

# """Получаем доступ к активной книге"""
# wb = Excel.ActiveWorkbook

# """Получаем доступ к активному листу"""
# sheet = wb.ActiveSheet

# """Формулы"""
# # sheet.Range("A5").FormulaR1C1 = "=R[-3]C+R[-3]C[1]"
# # sheet.Range("A5").Formula = "=A2+B2"

# cel = sheet.Cells(1, 1)
# # cel.Select
# cel.Copy()
# sheet.Cells(1, 3).Activate()
# sheet.Paste()




# def corect(sheet, StartNomerRow, EndNomerRow):
#     col = "F"
#     BBB = sheet.Range(f"{col}{StartNomerRow}:{col}{EndNomerRow}")
#     BBB.FormulaR1C1 = "=RC[-2]-RC[-1]"



# import win32com.client
# Excel = win32com.client.Dispatch("Excel.Application")
# Excel.Visible = 1
# # wb = Excel.Workbooks.Open(r"C:\Users\vvkhomutskiy\Desktop\test.xlsx")
# # wb = Excel.Workbooks.Open(r"C:\Users\vvkhomutskiy\Desktop\test.xlsx")
# wb = Excel.Workbooks.OpenDatabase(r"C:\Users\vvkhomutskiy\Desktop\test.xlsx")
# sheet = wb.Worksheets("Лист1")
# vals = sheet.Range(sheet.Cells(1, 1), sheet.Cells(2, 2))


# # vals = sheet.Cells(1, 1)
# # fo = vals.Formula
# # zn = [i.Font.Bold for i in vals]
# # print(f"data = {zn}")


# Excel = win32com.client.Dispatch("Excel.Application")
# wba = Excel.Workbooks.Add() 
# wba.Name = "test111"

# # copir = vals.Copy()
# # "test111.xlsx"

# wb = Excel.Workbooks.Open(r"C:\vxvproj\tnnc-Excel\collectorExcel\collectorApp\Сборная ведомость.xlsx")
# sheet = wb.Worksheets("Свод")

# sheet = wb.Worksheets
# sheet5 = wb.Sheets.Add().Name = "xxxxx"
# print(f"dir(Excel) = {dir(sheet)}")

# sheet.Range(sheet.Cells(10, 1), sheet.Cells(11, 2)).Activate()
# sheet.Paste()

# sheet.Range(sheet.Cells(15, 1), sheet.Cells(16, 2)).Activate()
# sheet.Paste()


# wb.Close()
# Excel.Quit()



# from time import sleep
# import win32com.client
# Excel = win32com.client.Dispatch("Excel.Application")

# # print(f"dir(Excel) = {dir(Excel)}")
# xxx = Excel.Workbooks
# # print(f"xxx = {dir(xxx)}")

# xxx = Excel.Workbooks.Open(r"C:\Users\vvkhomutskiy\Desktop\test.xlsx").Close()
# # xxx = Excel.Workbooks.Open
# print(f"xxx = {dir(xxx)}")
# # wb.Activate()
# # sleep(3.0)
# # wb.Close()
# # Excel.Quit()



import win32com.client
Excel = win32com.client.Dispatch("Excel.Application")
# wb = Excel.Workbooks.Open(r"C:\Users\vvkhomutskiy\Desktop\test.xlsx")
wb = Excel.ActiveWorkbook
# wb = Excel.Workbooks.Creator()
# xxx = wb.Workbooks
# xxx = wb
# Excel.Saved(False)


# wb.Close.Savechanges = False
# wb.Close(False)
wb.Close(False)

# wb.close
# print(f"xxx = {xxx}")
# print(f"xxx = {dir(xxx)}")


# Excel.Visible = 1
# wb = Excel.Workbooks.Open(r"C:\Users\vvkhomutskiy\Desktop\test.xlsx")
# wb = Excel.Workbooks.Open(r"C:\Users\vvkhomutskiy\Desktop\test.xlsx")



# FileName = r"C:\Users\vvkhomutskiy\Desktop\test.xlsx"
# CommandText = 
# CommandType = 
# BackgroundQuery = 
# ImportDataAs = 
# wb = Excel.Workbooks.OpenDatabase (FileName, CommandText, CommandType, BackgroundQuery, ImportDataAs)


# wb = Excel.Workbooks.OpenDatabase(r"C:\Users\vvkhomutskiy\Desktop\test.xlsx")