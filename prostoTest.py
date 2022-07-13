from rich import print
# aaa = 5
# bbb = 9
# ccc = 15
# if aaa == 1 or bbb == 1 or ccc == 15:
#     print("+++++++++++++++")
# ddd , lll = 555, 888

# print(ddd)
# print(lll)

# import numpy as np
# aaa = [0.253, 0.256, 0.244, 0.261, 0.244, 0.319, 0.264, 0.277, 0.273, 0.256]
# lll = np.std(aaa, ddof = 1.0)
# print(lll)


# aaa = [[111], [222, 333]]
# xxx = [aaa[g] + aaa[g + 1] for g in range(0, len(aaa) - 1)]
# sss = aaa[0] + aaa[1]
# print(f"sss = {sss}")
# print(f"xxx = {xxx}")

# print(f"sum(aaa) = {sum(aaa)}")

# prom = []
# prom = tuple([None] * 7)
# print(f"prom = {prom}")
# prom = None, None, None, None, None, None, None
# print(f"prom000 = {prom}")


# dict = {}
# for x in codeNomer:
#     xxx = []
#     for i in datasort:
#         if i[-1] == x:
#             xxx.append(i)
#     dict[x] = xxx

# import numpy as np
# # sssssss = 0.196339434276206
# sssssss = 0.44, 0.428140097


# dddd = np.std(sssssss, ddof = 1.0)
# print(f"0000000000000 = {dddd}")


# import sys, traceback, os
# from PyQt5 import QtCore, QtWidgets
# # os.system('CLS')

# app = QtWidgets.QApplication(sys.argv)

# def SMS(Text):
#     QtWidgets.QMessageBox.information(QtWidgets.QWidget(), 'Ошибка', Text)

# def GO():
#     try:
#         a = 5 / 0
#         print(f"a = {a}")
#     except:
#         SMS(traceback.format_exc())


# if __name__ == "__main__":
#     GO()
#     # SMS("Проверка")
#     # app = QtWidgets.QApplication(sys.argv)
#     sys.exit(app.exec_())    


# mygenerator = (x*x for x in range(3))
# mygenerator = [x*x for x in range(3)]
# print(f"mygenerator = {mygenerator}")
# for i in mygenerator :
#     print(i)
# for i in mygenerator :
#     print(i)

aaa = [1,3,5,7]

def aaafff(aaa):
    yield from aaa
        # print(i)
# bbb = [8,9,10,11]
ddd = aaafff(aaa)
# for b in ddd:
#     print(b)
print(next(ddd))
print(next(ddd))
# print(next(ddd))
# print(next(ddd))
