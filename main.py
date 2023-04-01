# импортируем необходимые библиотеки
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border, Side

# чтение файла Источник.XLS и создание временной таблицы istochnik
istochnik = pd.read_excel("Источник.XLS", usecols=[0, 1, 3, 4, 6], names=["modul", "kanal", "tipe", "signal", "koment"])

# создание файла Результат.xlsx
wb = Workbook()
wb.save("Результат.xlsx")

# получение уникальных значений из столбца modul
modul_values = istochnik["modul"].unique()

# запись таблиц в файл Результат.xlsx
with pd.ExcelWriter("Результат.xlsx") as writer:
    for modul in modul_values:
        # фильтрация таблицы istochnik по значению modul
        filtered_table = istochnik.loc[istochnik["modul"] == modul, ["kanal", "tipe", "signal", "koment"]]
        # запись таблицы в файл Результат.xlsx с именем листа modul
        filtered_table.to_excel(writer, sheet_name=modul, header=False, index=False)
        # получение последней записи (ячейки) в таблице и ее координат
        last_cell = writer.sheets[modul].cell(row=writer.sheets[modul].max_row, column=writer.sheets[modul].max_column)
        # создание стиля для обводки таблицы жирной линией
        border_style = Border(left=Side(style='thick'),
                              right=Side(style='thick'),
                              top=Side(style='thick'),
                              bottom=Side(style='thick'))
        # применение стиля ко всем ячейкам в таблице
        for row in writer.sheets[modul].rows:
            for cell in row:
                cell.border = border_style

print("Готово!")
