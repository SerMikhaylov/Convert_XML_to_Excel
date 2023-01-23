# импортируем необходимые библиотеки
import pandas as pd
import xml.etree.ElementTree as ET

# tree = ET.parse('Test_source.xml'.strip())
tree = ET.parse('data_Source_1.xml'.strip())
root = tree.getroot()
col_name = []
context = []
for child in root.iter():
    col_name.append(child.tag)
    context.append(child.text)
column_names = []  # список всех столбцов для включения в dataframe
for i in col_name:
    if i not in column_names:
        column_names.append(i)
length_column_names = len(column_names)
print("[INFO] lenght_column_name = ", length_column_names)

# добавим отсутствующие элементы, чтобы в каждом столбце dataframe было одинаковое число элементов
for ind_col in range(length_column_names):
    for ind_num in range(len(col_name)):
        if column_names[ind_col] == col_name[ind_num] and ind_col == ind_num % (length_column_names):
            continue
        elif column_names[ind_col] != col_name[ind_num] and ind_col == ind_num % (length_column_names):
            for ind_column_names in range(length_column_names):
                if col_name[ind_num] == column_names[ind_column_names] and ind_num % (length_column_names) > ind_column_names:
                    s = 0
                    for elem_1 in range(ind_num % (length_column_names), length_column_names):
                        col_name.insert(ind_num + s, column_names[elem_1])
                        context.insert(ind_num + s, None)
                        s += 1
                    i = 0
                    for elem_2 in range(ind_column_names):
                        col_name.insert(ind_num + s + i, column_names[elem_2])
                        context.insert(ind_num + s + i, None)
                        i += 1

                elif col_name[ind_num] == column_names[ind_column_names] and ind_num % (length_column_names) < ind_column_names:
                    i = 0
                    for elem in range(ind_col, ind_column_names):
                        col_name.insert(ind_num + i, column_names[elem])
                        context.insert(ind_num + i, None)
                        i += 1
                else:
                    continue
        elif column_names[ind_col] != col_name[ind_num] and ind_col != ind_num % (length_column_names):
            for num_column_names in range(length_column_names):
                if col_name[ind_num] == column_names[num_column_names] and ind_num % (length_column_names) == num_column_names:
                    continue
                elif col_name[ind_num] == column_names[num_column_names] and ind_num % (length_column_names) > num_column_names:
                    a = 0
                    for el in range(ind_num % (length_column_names), length_column_names):
                        col_name.insert(ind_num + a, column_names[el])
                        context.insert(ind_num + a, None)
                        a += 1
                    b = 0
                    for els in range(num_column_names):
                        col_name.insert(ind_num + a + b, column_names[els])
                        context.insert(ind_num + a + b, None)
                        b += 1
                elif col_name[ind_num] == column_names[num_column_names] and ind_num % (length_column_names) < num_column_names:
                    a = 0
                    for el in range(ind_num % (length_column_names), num_column_names):
                        col_name.insert(ind_num + a, column_names[el])
                        context.insert(ind_num + a, None)
                        a += 1
        else:
            continue
# проверка соответствия количества элементов в xml-файле и в названиях колонок будущего dataframe
sum_elem = len(col_name)
print("[INFO] последняя строка не заполнена до конца на количество элементов равных  = ", sum_elem % length_column_names)
if sum_elem % length_column_names != 0:
    for num_column_names in range(length_column_names):
        if col_name[sum_elem - 1] == column_names[num_column_names]:
            for element in range(num_column_names, length_column_names):
                try:
                    col_name.append(column_names[element + 1])
                    context.append(None)
                except IndexError:
                    col_name.append(column_names[element])
                    context.append(column_names[element])
        else:
            continue

# создадим словарь с ключем - названием колонок и значениями - элементами соответствующих тэгов
data = {}  # словарь для записи значений каждого элемента исходного массива данных
for ind_col in range(length_column_names):
    text_data = []  # создадим список для записи значений соответствующих элементов column_names
    for ind_num in range(len(col_name)):
        if column_names[ind_col] == col_name[ind_num] and ind_col == ind_num % (length_column_names):
            text_data.append(context[ind_num])
        data[column_names[ind_col]] = text_data

lenght = {}
for key, value in data.items():
    lenght[key] = len(value)
# print(lenght)

# сохранение данных в dataframe pandas
data_rezult = pd.DataFrame(data)
print('[INFO] DataFrame create successfull.')

# создаем объек для подготовки к записи данных в файл Excel
writer = pd.ExcelWriter('rezult.xlsx')

# записываем данные dataFrame в файл Excel
data_rezult.to_excel(writer)

# сохраняем файл Excel
writer.save()
print('[INFO] DataFrame is written successfully to Excel File.')