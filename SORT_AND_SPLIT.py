import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import filedialog


def chose_folder():
    print('Выберите папку для результата работы во всплывающем окне:')
    root = Tk()
    root.withdraw()
    root.focus_force()
    folder = filedialog.askdirectory(parent=root)
    folder = folder +'\\'
    return folder


def read_excel():
    print('Выберите файл для сортировки во всплывающем окне:')
    root = Tk()
    root.withdraw()
    root.focus_force()
    file = askopenfilename(parent=root)  # show an "Open" dialog box and return the path to the selected file
    dfs = pd.read_excel(file)
    print('Идет сортировка и подготовка массива')
    return dfs


def svod(list_svod, output_folder): # Формирование сводной таблицы с количеством объектов по каждому разделенному файлу
    out = pd.DataFrame(list_svod, columns=['SORT','AMOUNT'])
    svod_string = output_folder + 'svod.xlsx'
    out.to_excel(svod_string, header=['SORT','AMOUNT'], index=False)
    pass


def separate(dfs, output_folder, column_name):
    sort_value = dfs[column_name].to_list()
    sort_value = list(set(sort_value))
    list_sep = []
    for i in sort_value:
        try:
            DF_SORT = dfs.loc[dfs[column_name] == i]
            string = output_folder + str(i) + ".xlsx"
            DF_SORT.to_excel(string, index=False)
            list_sep.append([i, DF_SORT.shape[0]])
        except Exception:
            print('Возникла ошибка при записи объектов: ',i)
            list_sep.append([i, DF_SORT.shape[0]])
        continue
    return list_sep


if __name__ == "__main__":
    dfs = read_excel()
    output_folder = chose_folder()
    column_name = input('Введите уникальное наименование столбца для сортировки:  \n')
    list_sep = separate(dfs, output_folder, column_name)
    svod(list_sep, output_folder)
    print('Сортировка завершена, файлы помещены в эту папку: ',output_folder)
    print('Сводная таблица помещена в вышеуказанную папку')
    input('Нажмите ENTER чтобы закрыть программу')