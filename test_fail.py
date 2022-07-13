



# if __name__ == "__main__":
#     '''Открываем файл с кодировкой для чтения'''
#     f = open("data.txt", 'r', encoding='utf-8')
#     '''Читаем весь файл целиком как текст'''
#     # text = f.read()
#     '''Читаем файл и разбивает строку на подстроки в зависимости от разделителя'''
#     text = f.read().split("\n")

#     # text = f.read()
#     """Заменяем в строке одну подстроку на другую """
#     # text = text.replace("\n", " - ") 

#     print(text)

#     '''Обрезание пробелов слева и справа или "\\n" методом .strip() '''
#     """Чтение строк как список функцией readlines()"""
#     # lines = f.readlines()
#     '''Выбираем из списка строку нужную по индексу '''
#     # print(lines[0].strip())
#     """Альтернативный способ"""
#     # lines = [i.strip() for i in f.readlines()]
#     # print(lines)

#     """итерация по строкам"""
#     # for line in lines:
#     #     print(line.strip())
    
#     '''Закрываем файл'''
#     f.close()


# def setv():
#     with open("version.py", "r") as file:
#         text = float(file.read())
#     with open("version.py", "w") as file:
#         text += 0.01
#         text = round(text, 2)
#         file.write(str(text))

# if __name__ == "__main__":
#     setv()



# for filename in os.listdir(directory):
#     f = os.path.join(directory, filename)
#     # if os.path.isfile(f):
#     if os.path.isfile(f) and ".xls" in filename:
#         print(f, type(f))
#         os.startfile(f)
#         print("++++++++++++")



# for filename in os.scandir(directory):
#     print(filename, type(filename))

#     if filename.is_file():
#         print(filename.path)
#         # Book()
#         os.startfile(filename)
#         print("++++++++++++")

#         os.close(filename)