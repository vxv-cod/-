import os
os.system('CLS') 

# from docx import Document
# # создание документа
# # document = Document()

# # открытие документа
# doc = Document(r"C:\Users\vvkhomutskiy\Documents\vxv\code\ExPro\111.docx")
# # количество абзацев в документе
# lenn = len(doc.paragraphs)
# # print("lenn = ", lenn)

# # lenn = doc.paragraphs[5].text

# text = []
# for paragraph in doc.paragraphs:
#     text.append(paragraph.text)
# print('\n'.join(text))
# print(text)





# import win32com.client

# """Создаем COM объект"""
# wrd = win32com.client.Dispatch("Word.Application")
# # wrd.Visible = False
# Doc = wrd.Documents.Open(r"C:\Users\vvkhomutskiy\Documents\vxv\code\ExPro\111.docx")
# """Получаем доступ к активному документу"""
# # activDoc = wrd.ActiveDocument

# # print("activDoc = ", activDoc)

# # fff = Doc.Range().Text
# # fff = Doc.Range().Paragraphs(2)
# fff = Doc.Paragraph[0]
# # fff = Doc.Range().Paragraphs(1).Range.Text
# print("fff = ", fff)


# Doc.Close()
# wrd.Quit(0)

# Sub vxv_normal_modul()
# With ActiveDocument.Range.Find
# '    Óáèðàåì ðàçðûâ ñòðàíèöû
# '    .Text = "^m"
# '    .Execute Replace:=wdReplaceAll
# '    Óáèðàåì ðàçðûâ ðàçäåëà
#     .Text = "^b"
#     .Execute Replace:=wdReplaceAll
# '    Óáèðàåì àáçàöû(enter)
# '    .Text = "^p"
# '    .Execute Replace:=wdReplaceAll
    
#     MsgBox "Ìàêðîñ <vxv_normal_modul> îòðàáîòàë"
# End With

# End Sub



import win32com.client

"""Создаем COM объект"""
wrd = win32com.client.Dispatch("Word.Application")
# wrd.Visible = False
# Doc = wrd.Documents.Open(r"C:\Users\vvkhomutskiy\Documents\vxv\code\ExPro\111.docx")
# """Получаем доступ к активному документу"""
ad = wrd.ActiveDocument
tab1 = ad.Tables(3)

ColumnsCount = tab1.Columns.Count
RowsCount = tab1.Rows.Count
# xxx = tab1.Columns(2)
# xxx = tab1.Cell(2, 1)
Col1 = tab1.Columns(1).Cells
Col2 = tab1.Columns(2).Cells
# xxx = len(Col1)
# xxx = Col1.Count
# xxx = [str(Col1[i]).split("\r")[0] for i in range(len(Col1))]
# print(f"xxx = {xxx}")
# xxx = [str(Col2[i]).split("\r")[0] for i in range(len(Col2))]
# xxx = [str(Col2[i]).split("\r")[:-1] for i in range(len(Col2))]
# print(f"xxx = {xxx}")
for i in range(len(Col2)):
    xxx = str(Col2[i]).split("\r")[:-1]
    xxx = ' '.join(xxx)
    # xxx = str(Col2[i])
    # xxx = Col2[i]
    print(f"xxx = {xxx}")

xxx = [str(Col2[7])], [str(Col2[8])]
print(f"xxx = {xxx}")

xxx = tab1.Cell(8, 2)
xxx = str(xxx)
print(f"xxx = {xxx}")
xxx = str(xxx).split("\r")[:-1]
xxx = ' '.join(xxx)
print(f"xxx = {xxx}")


# tabCount = ad.Tables.Count
# for i in range(1, tabCount + 1):
#     xxx = ad.Tables(i)
#     print(f"tab {i} = {xxx}")
