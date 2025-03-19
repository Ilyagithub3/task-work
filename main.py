import os
from os import path
import tkinter as tk
from tkinter import scrolledtext, font, messagebox, filedialog
import itertools
import re
import docx
def mixed_word(word):
    '''Эта функция принимает на вход слово word и возвращает список всех возможных комбинаций букв, которые могут быть созданы из букв, найденных в словаре dictionary.txt для каждой буквы из введенного слова'''
    # Открытие файла словаря
    str(word).lower()
    os.chdir('C:\\Users\Илья\\PycharmProjects\\Парсинг файлов')
    f = open("dictionary.txt", encoding="UTF-8")
    b = list()
    # Чтение файла построчно
    for i in f:
        b.append([])
        # Разбиение строки на символы и добавление в список
        for j in i[:-1]:
            b[-1].append(j)
    c = list()
    # Поиск всех комбинаций букв, соответствующих введенному слову
    for i in word:
        for j in b:
            if i in j:
                c.append(j)
    # Генерация всех возможных комбинаций букв
    d = list(map(''.join, itertools.product(*c)))
    return d

#Функция для создания мануала
def manual(event):
    manul = tk.Tk()
    manul.title("Мануал")
    manul.geometry("900x700")
    # Создание виджета Text для отображения мануала
    text_widget = tk.Text(manul, font=('Arial', 12), wrap='word')
    text_widget.pack(expand=True, fill='both')

    # Текст мануала
    manual_text = """
    # Руководство пользователя для программы "Парсинг файлов"

  Главное окно программы содержит следующие элементы:

    ##Поле "Путь до директории":  Сюда вводится путь к директории, которую необходимо проанализировать.  Можно ввести вручную или выбрать через кнопку "Обзор".

    ##Кнопка "Обзор": Открывает диалоговое окно для выбора директории.

    ##Поле "Слово для поиска":  Сюда вводится слово, которое нужно найти в файлах (только для функционала поиска слов в файлах).

    ##Кнопка "Все папки и файлы":  Выводит в область результатов список всех папок и файлов в выбранной директории.

    ##Кнопка "Поиск папки/файла": Выводит в область результатов найденный файл или папку по заданному слову.

    ##Кнопка "Поиск слова": Ищет указанное слово (и его анаграммы) в .txt и .docx файлах выбранной директории и выводит результаты в область результатов.

    """

    # Вставка текста в виджет Text
    text_widget.insert('1.0', manual_text)

    # Запрет на редактирование текста
    text_widget.config(state='disabled')

    manul.mainloop()

def vb_dir():
    dirpath = filedialog.askdirectory()
    var.set(dirpath)


def output_scrolled():
    global output, is_output_scrolled
    direct = entry_dir.get()
    if is_output_scrolled and entry_dir.get() != '' and os.path.exists(direct) == True:
        is_output_scrolled = False
        output = scrolledtext.ScrolledText(root, height=36, width=128)
        output.pack()
        output.tag_config("found", foreground="#191970")
        output.tag_config("count", foreground="green")
        output.tag_config("papka", foreground = "#D2691E")
        output.tag_config("not_found", foreground="red")

def full_papka_and_files():
    direct = entry_dir.get()
    if entry_dir.get() != '':
        if os.path.exists(direct) == True:
            output.delete('1.0', tk.END)
            os.chdir(direct)
            address_list = []
            dirs_list = []
            files_list = []
            for address, dirs, files in os.walk(direct):
                address_list.append(address)
                for i in dirs:
                    dirs_list.append(i)
                for j in files:
                    files_list.append(j)
            address_full = []
            for i in address_list:
                for j in dirs_list:
                    absol = path.join(i, j)
                    if path.isdir(absol):
                        address_full.append(absol)
            if not dirs_list:
                output.insert(tk.END, "Папок нет!\n\n", 'not_found')
            else:
                output.insert(tk.END, "Все папки:\n", 'found')
                for i, j in zip(dirs_list, address_full):
                    output.insert(tk.END, f"Название: '{i}'. ", "papka")
                    output.insert(tk.END, f"Путь: '{j}'\n")
                output.insert(tk.END, f"Количество папок: {len(dirs_list)}\n\n", "count")
            full_filles = []
            for i in address_list:
                for j in files_list:
                    absol1 = path.join(i, j)
                    if path.isfile(absol1):
                        full_filles.append(absol1)
            if not files_list:
                output.insert(tk.END, "Файлов нет!\n\n", 'not_found')
            else:
                output.insert(tk.END, "Все файлы:\n", 'found')
                for i, j in zip(files_list, full_filles):
                    output.insert(tk.END, f"Название: '{i}'. ", "papka")
                    output.insert(tk.END, f"Путь: '{j}'\n")
                output.insert(tk.END, f"Количество файлов: {len(files_list)}\n\n", "count")
        else:
            messagebox.showwarning(title='Предупреждение', message=f"Директории с названием '{direct}' не существует! Попробуйте снова!")
    else:
        messagebox.showwarning(title='Предупреждение', message="Ошибка: Введите путь директории")
def find_files_papka():
    direct = entry_dir.get()
    file_papka = entry_word.get()
    mixed = mixed_word(file_papka)
    if entry_dir.get() != '':
        if os.path.exists(direct) == True:
            if entry_word.get() != '':
                output.delete('1.0', tk.END)
                os.chdir(direct)
                address_list = []
                dirs_list = []
                files_list = []
                for address, dirs, files in os.walk(direct):
                    address_list.append(address)
                    for i in dirs:
                        dirs_list.append(i)
                    for j in files:
                        files_list.append(j)
                address_full = []
                for i in address_list:
                    for j in dirs_list:
                        absol = path.join(i, j)
                        if path.isdir(absol):
                            address_full.append(absol)
                full_filles = []
                proverka = True
                for i in address_list:
                    for j in files_list:
                        absol1 = path.join(i, j)
                        if path.isfile(absol1):
                            full_filles.append(absol1)

                for i in files_list:
                    for j in mixed:
                        if j in i:
                            proverka = False
                            ind = files_list.index(i)
                            output.insert(tk.END, f"Файл с именем '{file_papka}' найден. ", 'papka')
                            output.insert(tk.END, f"Путь '{full_filles[ind]}'\n")
                for i in dirs_list:
                    for j in mixed:
                        if j in i:
                            proverka = False
                            ind = dirs_list.index(i)
                            output.insert(tk.END, f"Папка с именем '{file_papka}' найдена. ", 'papka')
                            output.insert(tk.END, f"Путь '{address_full[ind]}'\n")
                if proverka:
                    output.insert(tk.END, f"Файл или папка с именем '{file_papka}' отсутствуют. ", 'not_found')
            else:
                messagebox.showwarning(title='Предупреждение', message=f"Ошибка: Введите название файла или папки")
        else:
            messagebox.showwarning(title='Предупреждение', message=f"Директории с названием '{direct}' не существует! Попробуйте снова!")
    else:
        messagebox.showwarning(title='Предупреждение', message="Ошибка: Введите путь директории")

def find_word():
    direct = entry_dir.get()
    word = entry_word.get()
    mixed = mixed_word(word)
    if entry_dir.get() != '':
        if os.path.exists(direct) == True:
            if entry_word.get() != '':
                output.delete('1.0', tk.END)
                os.chdir(direct)
                address_list = []
                dirs_list = []
                files_list = []
                for address, dirs, files in os.walk(direct):
                    address_list.append(address)
                    for i in dirs:
                        dirs_list.append(i)
                    for j in files:
                        files_list.append(j)
                full_filles = []
                for i in address_list:
                    for j in files_list:
                        absol1 = path.join(i, j)
                        if path.isfile(absol1):
                            full_filles.append(absol1)
                proverka = True
                files_txt = []
                files_docx = []
                files_xlsx = []
                for i in full_filles:
                    filename, file_extension = os.path.splitext(i)
                    if file_extension == '.txt':
                        files_txt.append(i)
                    if file_extension == '.docx':
                        files_docx.append(i)
                    if file_extension == '.xlsx':
                        files_xlsx.append(i)
                # Поиск слов в .txt файлах и указание номера строки
                for i in files_txt:
                    with open(i, 'r', encoding='utf-8') as f:
                        # Чтение файла построчно
                        for line_number, line in enumerate(f, 1):
                            words = re.findall(r'\b\w+\b', line)
                            for j in words:
                                for k in mixed:
                                    if k == j:
                                        for h in files_list:
                                            if h in i:
                                                proverka = False
                                                # Вывод информации о найденном слове и номере строки
                                                output.insert(tk.END,
                                                              f"Слово '{k}' найдено в файле '{h}' в строке {line_number}\n",
                                                              'papka')
                                                output.insert(tk.END, f"Путь '{i}'\n")
                for i in files_docx:
                    doc = docx.Document(i)
                    result = [p.text for p in doc.paragraphs]
                    for j in result:
                        words = re.findall(r'\b\w+\b', j)
                        for k in words:
                            for o in mixed:
                                if k == o:
                                    for h in files_list:
                                        if h in i:
                                            proverka = False
                                            output.insert(tk.END, "\n\n")
                                            output.insert(tk.END,
                                                          f"Слово '{k}' найдено в файле '{h}'\n",
                                                          'papka')
                                            output.insert(tk.END, f"Путь '{i}'\n")
                if proverka:
                    output.insert(tk.END, f"Слово '{word}' не найдено в файлах. ", 'not_found')
            else:
                messagebox.showwarning(title='Предупреждение', message=f"Ошибка: Введите слово!")
        else:
            messagebox.showwarning(title='Предупреждение', message=f"Директории с названием '{direct}' не существует! Попробуйте снова!")
    else:
        messagebox.showwarning(title='Предупреждение', message="Ошибка: Введите путь директории")

root = tk.Tk()
root.title("Парсинг файлов и папок")
root.state('zoomed')

is_output_scrolled = True
is_output_scrolled1 = False
is_vb_dir = False


font3 = font.Font(weight = 'bold', size = 8)
lb3 = tk.Label(root, text = "Мануал F1", font = font3)
lb3.pack(anchor = 'nw')

font1 = font.Font(weight = 'bold', size = 9)
lb1 = tk.Label(root, text = 'Введите путь директории', font = font1)
lb1.pack()

top_dir = tk.Frame(root)
top_dir.pack(padx = 6, ipadx = 10)

btn4 = tk.Button(top_dir, text = "Выбрать директорию",  bd = 4, activebackground = 'red', activeforeground = 'yellow', command = vb_dir)
btn4.pack(side = 'right')
var = tk.StringVar()
entry_dir = tk.Entry(top_dir, width = 50, bd = 3, textvariable = var)
entry_dir.pack(padx = (144, 20))



btn1 = tk.Button(root, text = "Вывести все файлы и папки", command = lambda:[output_scrolled(), full_papka_and_files()], bd = 4, activebackground = 'red', activeforeground = 'yellow' )
btn1.pack(pady = 6)

font2 = font.Font(weight = 'bold', size = 9)
lb2 = tk.Label(root, text = 'Введите слово, папку или файл', font =font2)
lb2.pack(pady = (15, 0))


entry_word = tk.Entry(root, width = 50, bd = 3)
entry_word.pack()

f_top = tk.Frame(root)
f_top.pack(pady = 6, ipadx = 10)

# Создание и размещение кнопок внутри фрейма
btn2 = tk.Button(f_top, text="Найти слово", width=15, bd = 4, activebackground = 'red', activeforeground = 'yellow', command = lambda:[output_scrolled(), find_word()])
btn2.pack(side = 'left')

btn3 = tk.Button(f_top, text="Найти папку, файл", width=15, bd = 4, activebackground = 'red', activeforeground = 'yellow', command = lambda:[output_scrolled(), find_files_papka()])
btn3.pack(side = 'right')

root.bind('<Control-F1>', manual)

tk.mainloop()
