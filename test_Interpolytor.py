'''Модуль двойно интерполяции'''
'''Функция нахождения индексов смежных колонок с искомым значением'''
def index_search(t1, x):
    for i in range(len(t1)):
        if t1[0] < t1[-1]:
            if x > t1[i]: continue
            else: break
        if t1[0] > t1[-1]:
            if x < t1[i]: continue
            else: break
    return i, abs(i-1)

def interpoi(t1, t2, yy, lineEditEnd):
    print("-------------------------------------")
    # lineEditEnd.setStyleSheet("background-color: rgb(220, 220, 220);")
    i = index_search (t1, yy)
    a1 = t1[i[0]]; a2 = t1[i[1]]; b1 = t2[i[0]]; b2 = t2[i[1]] # значения индексов
    if yy == a1:
        if isinstance(b1, str) and isinstance(b2, str):
            return "-"
        if isinstance(b1, float) and isinstance(b2, str):
            return b1
        if isinstance(b1, str) and isinstance(b2, float):
            # lineEditEnd.setStyleSheet("background-color: rgb(255, 170, 0);")
            return b2
    if (isinstance(b1, float) and isinstance(b2, str)):
        # lineEditEnd.setStyleSheet("background-color: rgb(255, 170, 0);")
        try:
            b2 = float(b2)
        except:
            b2 = b1
    if (isinstance(b1, str) and isinstance(b2, float)):
        # lineEditEnd.setStyleSheet("background-color: rgb(255, 170, 0);")
        try:
            b1 = float(b1)
        except:
            b1 = b2
    try:
        # y = round(b2 + ((yy - a2) * (b1 - b2) / (a1 - a2)), 4)
        y = round(b2 + ((yy - a2) * (b1 - b2) / (a1 - a2)), 4)
    except:
        y = "-"  
    print(f"y = b2:{b2} + ((yy:{yy} - a2:{a2})|{yy - a2}| * abs(b1:{b1} - b2:{b2})|{(b1 - b2)}| / (a1:{a1} - a2:{a2})|{(a1 - a2)}| = {y}")
    return y

def GO(RowKey, columnKey, RValue, YOOY, XOOX, lineEditEnd):
    # lineEditEnd.setStyleSheet("background-color: rgb(220, 220, 220);")
    dct = {RowKey[i] : RValue[i] for i in range(len(RValue))}
    #===============================================================================
    '''Находим индексы и значения строк'''
    indexRow = index_search(RowKey, YOOY)
    Y1 = RowKey[indexRow[0]]
    Y2 = RowKey[indexRow[1]]
    #===============================================================================
    '''Находим промежуточные значение в строке по столбцам'''
    XY1 = interpoi(columnKey, dct[Y1], XOOX, lineEditEnd)
    XY2 = interpoi(columnKey, dct[Y2], XOOX, lineEditEnd)
    #===============================================================================
    '''Интерполяция между найденными значениями по строками'''
    res = round(XY1 - ((Y1 - YOOY) * abs(XY1 - XY2) / abs(Y1 - Y2)), 4)
    print(f"res = XY1: {XY1} - ((Y1: {Y1} - YOOY: {YOOY}) * abs(XY1: {XY1} - XY2: {XY2}) / abs(Y1: {Y1} - Y2: {Y2}) = {res}")
    return str(res)
    #===============================================================================
if __name__ == "__main__":
    import os
    os.system('CLS')
    Value_Sypesi_Nezasol_Yth = [
                  # [0.00, 1.00, 2.00, 3.00, 4.00, 5.00, 6.00, 7.00, 8.00, 9.00, 10.0]   # index     
                  # [2.00, 1.00, 0.60, 0.40, 0.35, 0.30, 0.25, 0.20, 0.15, 0.10, 0.05]      
                    ["--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--"],  # 0.4
                    ["--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--"],  # 0.7
                    ["--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--"],  # 1.0
                    ["--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--"],  # 1.2
                    ["--", "--", "--", "--", 1.80, 1.74, 1.57, 1.33, 1.10, 0.93, 0.64],  # 1.4
                    ["--", "--", "--", "--", "--", "--", 1.80, 1.62, 1.45, 1.16, 0.81],  # 1.6
                    ["--", "--", "--", "--", "--", "--", "--", 1.86, 1.68, 1.45, 0.98],  # 1.8
                    ["--", "--", "--", "--", "--", "--", "--", "--", "--", "--", "--"]]  # 2.0

    RowKey_2 = [0.4, 0.7, 1.0, 1.2, 1.4, 1.6, 1.8, 2.0]
    columnKey_netor = [2.00, 1.00, 0.60, 0.40, 0.35, 0.30, 0.25, 0.20, 0.15, 0.10, 0.05]
    Ro = 1.58
    Wtot = 0.286

    GO(RowKey_2, columnKey_netor, Value_Sypesi_Nezasol_Yth, Ro, Wtot, None)