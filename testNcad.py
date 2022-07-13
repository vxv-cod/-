    
import win32com.client
# from pythoncom import VT_ARRAY, VT_R8, VT_DISPATCH
# from pythoncom import CoInitializeEx as pythoncomCoInitializeEx

# pythoncomCoInitializeEx(0)
acad = win32com.client.Dispatch("NanoCAD.Application")
doc = acad.ActiveDocument
msp = doc.ModelSpace

countelem = msp.Count
print(('countelem = ', countelem))

# doc.SetVariable('DIMTIH', "1")
# doc.SetVariable('DIMTAD', '1')
for object in msp:
    name = object.objectName
    # object.TextOutsideAlign = True
    # object.TextInsideAlign = False
    # object.VerticalTextPosition = 1
    '''
    1 - на линией
    2 - снаружи
    4 - под линией
    '''
    object.VerticalTextPosition = 1
    # object.HorizontalTextPosition = 0
    # print(f"name = {name}")
    # if name == "AcDbAlignedDimension":
    # if name == "AcDbRotatedDimension":
    #     print(f"name = {name}")
    #     object.VerticalTextPosition = 4
        # object.HorizontalTextPosition = 0

        # xxx = doc.GetVariable('DIMTAD')
        # print(f"DIMTAD = {doc.GetVariable('DIMTAD')}")
        # doc.SetVariable('DIMTAD', '1')
        # object.Footer()
        # xxx = object.TextOutsideAlign = True
        # object.HorizontalTextPosition = 0
        # object.VerticalTextPosition = 1
        # print(f"xxx = {xxx}")

# print(f"DIMTIH = {doc.GetVariable('DIMTIH')}")
# print(f"DIMTAD = {doc.GetVariable('DIMTAD')}")
# print(f"DIMGAP = {doc.GetVariable('DIMGAP')}")



print("====================================")