# Генератор списков с условием "if"
lst =  [1,5,3,7,3,4,8,3,9,3,1,2,8,6,7,4,9]
lst2 = [i for i in lst if i >= 5]
# Генератор списков с условием "if / else"
lst2 = [i + 5 if i < 5 else i for i in lst]

# Сортировка вложенных списков по 1у и 2м ключам
# 1 ключ
response = sorted(response, key = lambda i: i['StageNumber'])
# 2 ключа
# Создаем новый список или перезаписываем старый
response = sorted(response, key = lambda i: (i['StageNumber'], i['GenplanNumber']))
# Сортируем существующий список, изменяя его
response.sort(key = lambda item: (item['StageNumber'], item['GenplanNumber']))

'''
Функция enumerate() вернет кортеж, содержащий отсчет от start и значение
'''
for i, val in enumerate(lst):
    print(f'№ {i} => {val}')

for i, val in enumerate(lst, start=1):
    print(f'№ {i} => {val}')

'''
Получение списка парных кортежей (number, value) 
(порядковый номер в последовательности, значение последовательности)
'''
seasons = ['Spring', 'Summer', 'Fall', 'Winter']
list(enumerate(seasons))
[(0, 'Spring'), (1, 'Summer'), (2, 'Fall'), (3, 'Winter')]

# можно указать с какой цифры начинать считать
list(enumerate(seasons, start=1))
[(1, 'Spring'), (2, 'Summer'), (3, 'Fall'), (4, 'Winter')]


'''----------------------------------------------------------------------
Использование enumerate() для нахождения индексов 
минимального и максимального значений в числовой последовательности:
'''
lst = [5, 3, 1, 0, 9, 7]
# пронумеруем список 
lst_num = list(enumerate(lst, 0))
# получился список кортежей, в которых 
# первый элемент - это индекс значения списка, 
# а второй элемент - само значение списка
lst_num
# [(0, 5), (1, 3), (2, 1), (3, 0), (4, 9), (5, 7)]

# найдем максимум (из второго значения кортежей)
tup_max = max(lst_num, key=lambda i : i[1])
tup_max
# (4, 9)
f'Индекс максимума: {tup_max[0]}, Max число {tup_max[1]}'
# 'Индекс максимума: 4, Max число 9'

# найдем минимум (из второго значения кортежей)
tup_min = min(lst_num, key=lambda i : i[1])
tup_min
# (3, 0)
f'Индекс минимума: {tup_min[0]}, Min число {tup_min[1]}'
# 'Индекс минимума: 3, Min число 0'

'''----------------------------------------------------------------------'''