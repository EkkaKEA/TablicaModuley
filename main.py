import pandas as pd

# чтение файла с данными
df = pd.read_excel('Источник.xls', header=None, usecols=[0, 1, 3, 4, 6])

# временная таблица
istochnik = df.loc[:, ]
istochnik.columns = ['modul', 'kanal', 'tipe', 'signal', 'koment']

# разбиение таблицы по критерию в колонке 'modul'
tables = {}
for modul in istochnik['modul'].unique():
    tables[modul] = istochnik[istochnik['modul'] == modul]

# запись таблицы в новый файл
filename = 'Результат.xlsx'
writer = pd.ExcelWriter(filename)
for i, modul in enumerate(tables.keys()):
    # запись таблицы
    tables[modul].to_excel(writer, sheet_name=modul, index=False)

# добавление промежутков в 2 столбца
#workbook = writer.book
#worksheet = writer.sheets[f'Table {i + 1}']
#    for j in range(len(tables[modul].columns) * 2 - 1):
#worksheet.write(0, j + len(tables[modul].columns) + 1, '')

# добавление линии
#last_row = len(tables[modul].index) + 1
#format = workbook.add_format({'bottom': True})
#worksheet.conditional_format(f'A1:E{last_row}', {'type': 'formula', 'criteria': f'=$A1="{modul}"', 'format': format})
writer.save()
