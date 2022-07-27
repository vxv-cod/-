import os
from time import sleep
import win32com.client
import threading
from pythoncom import CoInitializeEx as pythoncomCoInitializeEx
from PyQt5 import QtCore, QtWidgets

from rich import print
from rich import inspect
# inspect(xxx, methods=True)
# inspect(xxx, all =True)
os.system('CLS')

class Signals(QtCore.QObject):
    '''
    sig.signal_Probar.emit(ui.progressBar_1, 100)
    sig.signal_label.emit(ui.label, "Выполнено")
    sig.signal_err.emit(f"Ошибка работы, повторите попытку \n\n{traceback.format_exc()}")
    sig.signal_color.emit(ui.progressBar_1, 0)
    sig.signal_color.emit(ui.progressBar_1, 1)
    sig.signal_bool.emit(True)
    sig.signal_bool.emit(ui.pushButton, False)
    '''
    signal_Probar = QtCore.pyqtSignal(list)
    signal_label = QtCore.pyqtSignal(list)
    signal_err = QtCore.pyqtSignal(str)
    signal_bool = QtCore.pyqtSignal(list)
    signal_color = QtCore.pyqtSignal(list)

    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)
        self.signal_Probar.connect(self.on_change_Probar,QtCore.Qt.QueuedConnection)
        self.signal_label.connect(self.on_change_label,QtCore.Qt.QueuedConnection)
        self.signal_err.connect(self.on_change_err,QtCore.Qt.QueuedConnection)
        self.signal_bool.connect(self.on_change_bool,QtCore.Qt.QueuedConnection)
        self.signal_color.connect(self.on_change_color,QtCore.Qt.QueuedConnection)

    '''Отправляем сигналы в элементы окна'''
    def on_change_Probar(self, s):
        '''Значение процента в прогресбаре'''
        s[0].setValue(s[1])
    def on_change_label(self, s):
        '''Отправляем текст в label'''
        s[0].setText(s[1])
    def on_change_err(self, s):
        '''Сообщение об ошибке'''
        eval(f"QtWidgets.QMessageBox.information(Form, 'Excel не отвечает...', {s})")
    def on_change_color(self, s):
        '''Устанавливаем цвет прогресбара'''
        if s[1] == 1:
            color = "170, 170, 170"
        else:
            color = "100, 150, 150"
        s[0].setStyleSheet("QProgressBar::chunk {background-color: rgb("f"{color}); margin: 2px;""}")
    def on_change_bool(self, s):
        s[0].setDisabled(s[1])

sig = Signals()


def thread(my_func):
    '''Обертка функции в потопк (декоратор)'''
    def wrapper():
        threading.Thread(target=my_func, daemon=True).start()
    return wrapper

import ctypes
def Allobject():
    '''Выясняем сколько объкетов Excel во всех экземлярах открыто'''
    EnumWindows = ctypes.windll.user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
    GetWindowText = ctypes.windll.user32.GetWindowTextW
    GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
    IsWindowVisible = ctypes.windll.user32.IsWindowVisible
    titles = []
    countExelList = []
    def foreach_window(hwnd, lParam):
        if IsWindowVisible(hwnd):
            length = GetWindowTextLength(hwnd)
            buff = ctypes.create_unicode_buffer(length + 1)
            GetWindowText(hwnd, buff, length + 1)
            titles.append((hwnd, buff.value))
        return True
    EnumWindows(EnumWindowsProc(foreach_window), 0)
    for i in range(len(titles)):
        if "- Excel" in  titles[i][1]:
            countExelList.append(1)
    countfail = sum(countExelList)
    return countfail

def ExcelInstances():
    '''Поиск всех процессов EXCEL.EXE'''
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    objSWbemServices = objWMIService.ConnectServer(".", "root\cimv2")
    colExcelInstances = objSWbemServices.ExecQuery(
            f"SELECT * FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
    return colExcelInstances

def decorExcel(my_func):
    '''Обертка функции в экземпляр Excel (декоратор)'''
    def wrapper():
        global objWorkbook, countfail
        countfail = Allobject()
        colExcelInstances = ExcelInstances()
        for objInstancei in colExcelInstances:
            objExcel = win32com.client.Dispatch("Excel.Application")
            for objWorkbook in objExcel.Workbooks:
                my_func()
            objExcel.Quit()
            objInstancei.Terminate
            sleep(2)
    return wrapper

'''Пример как исмользовать декоратор decorExcel'''
@decorExcel
def myfunc():
    WbName = objWorkbook.Name
    print(f"WbName = {WbName}")


def colorBar(s):
    '''Устанавливаем цвет прогресбара color = [ui.progressBar_1, 1]'''
    if s[1] == 1:
        color = "170, 170, 170"
    else:
        color = "100, 150, 150"
    s[0].setStyleSheet("QProgressBar::chunk {background-color: rgb("f"{color}); margin: 2px;""}")

def Book(fail=None, sheetName=None, ExcelVisible=1):
    '''Подключаемся к Excel'''
    pythoncomCoInitializeEx(0)
    Excel = win32com.client.Dispatch("Excel.Application")
    Excel.Visible = ExcelVisible
    if fail == None:
        '''Получаем доступ к активной книге'''
        wb = Excel.ActiveWorkbook
    else:
        '''Получаем доступ к определенному файлу'''
        wb = Excel.Workbooks.Open(rf"{fail}")

    if sheetName == None:
        """Получаем доступ к активному листу"""
        sheet = wb.ActiveSheet
    else:
        """Получаем доступ к определенному листу"""
        sheet = wb.Worksheets(sheetName)
    return wb, sheet

def StartEndCell(sheet):
    '''Определяем позиции первой и последней ячейки'''
    # '''Количество занимаемых таблицей строк'''
    count_row = sheet.UsedRange.Rows.Count
    # '''Количество занимаемых таблицей колонок'''
    count_col = sheet.UsedRange.Columns.Count
    # '''Номер первой занимаемой строчки'''
    StartRow = sheet.UsedRange.Row
    # '''Номер первой занимаемой колонки'''
    StartCol = sheet.UsedRange.Column
    # '''Номер последней занимаемой строчки'''
    EndRow = StartRow + count_row - 1
    # '''Номер последней занимаемой колонки'''
    EndCol = StartCol + count_col - 1
    return EndRow, EndCol

def RangeCells(sheet, StartRow, StartCol, EndRow, EndCol):
    '''Выделяем диапозон ячеек'''
    cell = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    return cell

def importdata(sheet, StartRow, StartCol, EndRow, EndCol):
    '''Собираем данные из диапозона ячеек'''
    cell = sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol))
    vals = cell.Formula
    if StartCol == EndCol:
        vals = [vals[i][x] for i in range(len(vals)) for x in range(len(vals[i]))]
    return vals

def exportdata(data, sheet, StartRow, StartCol, EndRow, EndCol):
    '''Отправляем данные в диапозон ячеек'''
    if StartCol == EndCol:
        data = [(i, None) for i in data]
        print(f"data = {data}")
    sheet.Range(sheet.Cells(StartRow, StartCol), sheet.Cells(EndRow, EndCol)).Formula = data

def grani(cel):
    '''Все грани тонкие в диапазоне'''
    cel.Borders.Weight = 2
    # cel.Borders.ColorIndex = 0
    # cel.Borders.Weight = 3          # Все границы
    # cel.Borders(1).Weight = 3       # Левая граница
    # cel.Borders(2).Weight = 3       # Правая граница
    # cel.Borders(3).Weight = 3       # Верхняя граница
    # cel.Borders(4).Weight = 3       # Нижняя граница
    # cel.Borders(5).Weight = 3       # Диагональ граница
    # cel.Borders(6).Weight = 3       # Диагональ граница
    # cel.Borders(7).Weight = 3       # Левая граница крайних ячеек в диапозоне
    # cel.Borders(8).Weight = 3       # Верхняя граница крайних ячеек в диапозоне
    # cel.Borders(9).Weight = 3       # Нижняя граница крайних ячеек в диапозоне
    # cel.Borders(10).Weight = 3      # Правая граница крайних ячеек в диапозоне
    # cel.Borders(11).Weight = 3      # Вертикальные внутренние границы ячеек в диапозоне
    # cel.Borders(12).Weight = 3      # Горизонтальные внутренние границы ячеек в диапозоне

def PatchFail(widgetText):
    strPath = str(widgetText)
    if "file:///" in strPath:
        strPath = strPath[8:]
    if strPath == '':
        sig.signal_err.emit(f"Не указана папка для сохранения файлов")
        return
    return strPath.replace("/", "\\")

def exportPDF(widgetText, objWorkbook):
    '''Экспорт в PDF'''
    strPath = PatchFail(widgetText)
    pdfName = objWorkbook.Name if ".xls" not in objWorkbook.Name else objWorkbook.Name.split(".xls")[0]
    # sheet.PrintOut(Copies=1, ActivePrinter="Microsoft Print to PDF", PrintToFile=True, PrToFileName = f"{strPath}\\{pdfName}.pdf")
    OutputFile = f"{strPath}\\{pdfName}.pdf"
    objWorkbook.ExportAsFixedFormat(0, OutputFile)










'''================================================================================================'''
'''================================================================================================'''
# inspect(cel.Font.Color, all=True)
# widgetText = ui.plainTextEdit.toPlainText()
# WbName = objWorkbook.Name
# wb, sheet = Book()

# EndRow, EndCol = StartEndCell(sheet)
# StartRow, StartCol = 1, 1

# cel = RangeCells(sheet, StartRow, StartCol, EndRow, EndCol)
# cel = sheet.Range("A5")
# cel = sheet.Rows(f"{StartRow}:{EndRow}")

'''Зачеркиваем текст в ячейке'''
# cel.Font.Strikethrough = True

'''Цвет текста в ячейке'''
# cel.Font.Color = 255

'''Сохранить как'''
# objWorkbook.SaveAs(f"{strPath}\\{objWorkbook.Name}{strFileExtension}", FileFormat=objWorkbook.FileFormat, CreateBackup=0)

'''Выбрать строчки'''
# RowsSelect = sheet.Rows(f"{StartRow}:{EndRow}")
'''Выбрать колонки'''
# ColSelect = sheet.Columns(f"{StartCol}:{EndCol}")

'''Объединение ячеек'''
# cel.Merge()
# cel.MergeCells = True

'''Перенести текст'''
# cel.WrapText = True

'''Задание ширины ячейки'''
# cel.ColumnWidth = 45
'''Задание высоты ячейки'''

'''Выравниваем текст в ячейке'''
'''Выравниваем по горизонтали (вертикали) центр'''
# cel.HorizontalAlignment = 3
# cel.VerticalAlignment = 2
'''Выравниваем по горизонтали влево'''
# cel.HorizontalAlignment = 1

'''Удаляем строки со сдвигом вверх'''
# sheet.Rows(f"{StartRow}:{EndRow}").Delete(1)

'''Удаляем колонки'''
# sheet.Columns("AV:AW").Delete(1)

'''Очистить содержимое строчек'''
# sheet.Rows(f"{StartRow}:{EndRow}").ClearContents()


'''Копируем ячейки'''
# cel = RangeCells(sheet, StartRow, StartCol, EndRow, EndCol)
# cel.Copy()

'''Вставить ячейки'''
# cel = RangeCells(sheet, StartRow, StartCol, EndRow, EndCol)
# cel.Activate()
# sheet.Paste()

"""Отключение уведомлений с ответом по умолчанию для сохранения без подтверждения"""
# Excel.DisplayAlerts = False  

'''Закрыть файл без сохранения'''
# wb.Close(False)
'''Закрыть файл с сохранением'''
# wb.Close()
'''Закрыть экземпляр Excel'''
# Excel.Quit()


# dirCell = dir(cel)
# for i in dirCell:
#     try:
#         xxx = f"cel.{i}()"
#         sss = eval(xxx)
#         print(f"{xxx} = {sss}")
#     except:
#         try:
#             xxx = f"cel.{i}"
#             sss = eval(xxx)
#             print(f"{xxx} = {sss}")
#             pass
#         except:
#             # print(f'///cel.{i} {type((i))} - не обработано................')
#             pass

