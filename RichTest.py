
'''https://rich.readthedocs.io/en/stable/introduction.html'''

from rich import print

data = [[206.0, 3.5, 0.253, 0.35, 0.177, 0.173, 0.439306358381503, 1.98, 2.7, 1.5802075019952113, 41.473796222399585, 0.708636363636364, 0.9639640795381651, 0.052, 
            None, None, None, None, None, None, None, None, None, None, None, 1.5, 1.74, 2.7, 2.33, 8.2, 2.34, 19.188, None, 0.23, 15.0, 19.0, None, None, None, None, 
            None, None, 1.3, 12.5, 36.4, 22.1, 27.7, 'Глина легкая пылеватая тугопластичная с примесью органического вещества', 1231.0],
        [206.0, 8.5, 0.256, 0.404, 0.201, 0.203, 0.27093596059113295, 1.94, 2.71, 1.5445859872611465, 43.00420711213482, 0.7545154639175258, 0.9194775099743128, 0.053, 
            None, None, None, None, None, None, None, None, None, None, None, 1.47, 1.72, 2.7, 2.34, None, None, None, None, None, None, None, None, None, None, None, 
            None, None, None, None, None, None, None, 'Глина тугопластичная с примесью органического вещества', 1031.0],
        [216.0, 3.0, 0.244, 0.362, 0.185, 0.177, 0.3333333333333333, 1.96, 2.7, 1.5755627009646302, 41.645825890198886, 0.7136734693877552, 0.923111238204175, 0.059, 
            None, None, None, None, None, None, None, None, None, None, None, 1.47, 1.71, 2.68, 2.31, 8.2, 2.34, 19.188, None, 0.25, 14.0, 22.0, None, None, None, 
            None, None, None, 1.5, 8.9, 39.3, 19.6, 30.7, 'Глина легкая пылеватая тугопластичная с примесью органического вещества', 1231.0]]

# print(f"data = {data}")
# mytable.add_rows(data)
# print(mytable)


'''Как выполнить расширенную проверку в Python'''
# from rich import inspect
# import rich
# inspect(rich)


'''Как оформить вашу консоль с помощью Rich'''
# from rich.console import Console
# console = Console()

# def merge_dict(dict_one, dict_two):
#     merged_dict = dict_one,  dict_two
#     console.log(merged_dict, log_locals=True)
# merge_dict({'id': 1}, {'name': 'Ashutosh'})

# console.log(f"{data}")


'''Как использовать Tree в Rich'''
# from rich.tree import Tree
# from rich import print as rprint
# tree = Tree("Family Tree")
# tree.add("[blue]Dad")
# fff = tree.add("[blue]Brother")
# fff.add("[blue]Son").add("[blue]Son")
# fff.add("[blue]Son")
# tree.add("[blue]Dad")
# rprint(tree)


'''Как отобразить индикатор выполнения с помощью Rich'''
# from rich.progress import track
# from time import sleep

# def process_data():
#     sleep(0.02)

# for _ in track(range(100), description='[green]Processing data'):
#     process_data()

'''Если мы хотим записать время завершения выполнения конкретной задачи, 
мы можем использовать вместо этого console.status.'''
# from rich.console import Console
# from time import sleep
# console = Console()
# data = [1, 2, 3, 4, 5]
# with console.status("[bold green]Fetching data...") as status:
#     while data:
#         num = data.pop(0)
#         sleep(1)
#         console.log(f"[green]Finish fetching data[/green] {num}")
#     console.log(f'[bold][red]Done!')


'''Вы можете работать напрямую с классом Progress, если вам нужно несколько задач на дисплее 
или вы хотите настроить столбцы на дисплее прогресса. После создания объекта Progress используйте 
( add_task()) для добавления задач и ( update_progress()) для обновления хода выполнения.
Класс Progress предназначен для использования в качестве менеджера контекста, 
автоматически запускающего и останавливающего отображение прогресса.'''

# import time
# from rich.progress import Progress
# with Progress() as progress:
#     task1 = progress.add_task("[red]Downloading...", total=100)
#     task2 = progress.add_task("[green]Processing...", total=100)
#     task3 = progress.add_task("[cyan]Installing...", total=100)
#     while not progress.finished:
#         progress.update(task1, advance=0.9)
#         progress.update(task2, advance=0.6)
#         progress.update(task3, advance=0.3)
#         time.sleep(0.02)

'''Как отображать таблицы в Python с помощью Rich'''
# from rich.console import Console
# from rich.table import Table
# table = Table(title="Todo List")
# table.add_column("S. No.", style="cyan", no_wrap=True)
# table.add_column("Task", style="magenta")
# table.add_column("Status", justify="right", style="green")
# table.add_row("1", "Buy Milk", "✅")
# table.add_row("2", "Buy Bread", "✅")
# table.add_row("3", "Buy Jam", "❌")
# console = Console()
# console.print(table)



# from rich import inspect
# from rich.color import Color
# color = Color.parse("red")
# inspect(color, methods=True)


'''https://pypi.org/project/prettytable/'''
# from prettytable import PrettyTable
# from prettytable import from_db_cursor
# mytable = PrettyTable()
# mytable.field_names = [i for i in range(0, 48 + 1)]

# print(f"data = {data}")
# mytable.add_rows(data)
# print(mytable)