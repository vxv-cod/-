import os
from time import sleep
import win32com.client
import threading
from pythoncom import CoInitializeEx as pythoncomCoInitializeEx
from PyQt5 import QtCore, QtWidgets
import sys
import traceback


# World = win32com.client.Dispatch("World.Application")

Word = win32com.client.Dispatch("Word.Application") 
Doc = Word.Documents.Open(r"C:\Users\vxv\Desktop\111.docx")
Doc.Tables(1).AutoFitBehavior(2)
Doc.Tables(1).Range.ParagraphFormat.Alignment = 1

'''Выравниваем по вертикали ячейку'''
Doc.Tables(1).Rows(4).Cells(5).VerticalAlignment = 3
'''Выравниваем по вертикали строку'''
Doc.Tables(1).Rows(4).Cells.VerticalAlignment = 3
'''Выравниваем по вертикали столбец'''
Doc.Tables(1).Columns(4).Cells.VerticalAlignment = 3
'''Выравниваем по вертикали всю таблицу'''
Doc.Tables(1).Range.Cells.VerticalAlignment = 3
