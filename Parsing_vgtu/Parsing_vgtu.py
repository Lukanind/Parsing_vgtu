# -*- coding: cp1251 -*-
import requests
import fitz  # PyMuPDF
from bs4 import BeautifulSoup
from bs4 import Tag, NavigableString

import re
from io import BytesIO
import pandas as pd
import os
import platform
from openpyxl import load_workbook

from tkinter import *
from tkinter import ttk

count_annot, count_rab, count_rpv, count_pp = 0, 0, 0, 0
a_link_list, r_link_list, v_link_list, p_link_list = [], [], [], []

def reading(url, ID):

    files = find_pdf(url, ID)
    name_list = []

    message.set("Проверяем файлики...")
    for pdf_file in files:  
        response = requests.get(pdf_file)
        pdf_content = response.content

        pdf_stream = BytesIO(pdf_content)
        document = fitz.open(stream=pdf_stream, filetype="pdf")
        try:
            #Поиск названия дисциплины
            text = ""
            for page_num in range(len(document)):
                page = document.load_page(page_num)
                text += page.get_text()
            if ID == 'rpv':
                pattern = re.compile(r'(рабочая)\s+(программа)\s+(воспитания)', re.IGNORECASE)
                matches = re.findall(pattern, text)
                result = [' '.join(match) for match in matches]
            else:
                pattern = re.compile(r'«([A-ZА-ЯЁ][^0-9«»]*)\s*([^0-9«»]*?)»', re.IGNORECASE)
                matches = re.findall(pattern, text)
                result = [''.join(match) for match in matches]
                filtered_result = [word for word in result if word and word[0].isupper() and not word[0].isdigit()]
                result = filtered_result
                for i in range(len(result)):
                    result[i] = re.sub("\n","",result[i])

            #print(result)
            #input()
            try:
                #Отбор неправильных данных 
                for res in result:
                        if res != "Воронежский государственный технический университет":
                            name_list.append(res)
                            break
            except:
                name_list.append(f"Не удалось прочитать файл: {pdf_file}")
                print(f"Не удалось прочитать файл: {pdf_file}")
        finally:
            document.close()  
    return name_list

def find_pdf(url, ID):
    error = False

    print("Начало парсинга...")
    message.set("Начало парсинга...")
    # id для аннотаций "arp", рабочих программ - "rp", раб прог воспитания - "rpv", программы практик - "pp"
    link_to_pdf = []
    #url = 'https://cchgeu.ru/education/programms/poas-3/?docs2021'  
    response = requests.get(url)

    soup = BeautifulSoup(response.text, "lxml")
    category = soup.find("h5", id = ID)
    
    for file in category.next_siblings:
        if file.name == "h5":
            break
        if isinstance(file, Tag):
            try:
                href = file.find("a", class_ = 'wb-ba').get('href')
                link_to_pdf.append('https://cchgeu.ru' + href)
            except:
                error = True
    
    param = "что-то"
    count = 0
    
    if ID == 'arp':
        param = "аннотаций"
        global count_annot
        global a_link_list
        count_annot = len(link_to_pdf)
        a_link_list = link_to_pdf
        count = count_annot
    elif ID == 'rp':
        param = "рабочих программ"
        global count_rab
        global r_link_list
        count_rab = len(link_to_pdf)
        r_link_list = link_to_pdf
        count = count_rab
    elif ID == 'rpv':
        param = "программ воспитания"
        global count_rpv
        global v_link_list
        count_rpv = len(link_to_pdf)
        v_link_list = link_to_pdf
        count = count_rpv
    elif ID == 'pp':
        param = "программ практик"
        global count_pp
        global p_link_list
        count_pp = len(link_to_pdf)
        p_link_list = link_to_pdf
        count = count_pp
    
    if (error):
        print("Возможно, здесь ничего нет 0_0")
    else:
        print(f"Было найдено {param}: {count}")

    message.set(f"Было найдено {param}: {count}")

    return link_to_pdf

def save_in_exel(list_annot, list_rab, list_rpv, list_pp, url):

    global count_annot, count_rab, count_pp, count_rpv
    global a_link_list, r_link_list, v_link_list, p_link_list

    count_annot_string = f"{count_annot}"
    count_rab_string = f"{count_rab}"
    count_rpv_string = f"{count_rpv}"
    count_pp_string = f"{count_pp}"

    df1 = pd.DataFrame(list_annot, columns=['Аннотации ' + "(" + str(count_annot) + ")"])
    df1.insert(1, "Ссылка", a_link_list)
    df2 = pd.DataFrame(list_rab, columns = ['Рабочие программы ' + "(" + str(count_rab) + ")"])
    df2.insert(1, "Ссылка", r_link_list)
    df3 = pd.DataFrame(list_rpv, columns=['Программы воспитания ' + "(" + str(count_rpv) + ")"])
    df3.insert(1, "Ссылка", v_link_list)
    df4 = pd.DataFrame(list_pp, columns=['Программы практик ' + "(" + str(count_pp) + ")"])
    df4.insert(1, "Ссылка", p_link_list)

    response = requests.get(url)
    soup = BeautifulSoup(response.text, "lxml")
    title = soup.find('a', class_ = 'active', string = re.compile('Нормативное'))

    filename = f'{title.text}.xlsx'
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='Аннотации', index=False)
        df2.to_excel(writer, sheet_name='Рабочие программы', index=False)
        df3.to_excel(writer, sheet_name='Программы воспитания', index=False)
        df4.to_excel(writer, sheet_name='Программы практик', index=False)

    print(f"Список предложений успешно сохранен в '{filename}'")

    filepath = os.path.abspath(filename)
    adjust_column_width(filepath, 'Аннотации')
    adjust_column_width(filepath, 'Рабочие программы')
    adjust_column_width(filepath, 'Программы воспитания')
    adjust_column_width(filepath, 'Программы практик')
    open_file(filepath)


def open_file(filepath):
    if platform.system() == 'Windows':
        os.startfile(filepath)
    elif platform.system() == 'Darwin':  # macOS
        os.system(f"open {filepath}")
    else:  # Linux 
        os.system(f"xdg-open {filepath}")

def adjust_column_width(filepath, sheet_name):
    wb = load_workbook(filepath)
    ws = wb[sheet_name]

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    wb.save(filepath)




if __name__ == "__main__":

    root = Tk()
    root.title("ПАРСЕР ОБРАЗОВАТЕЛЬНЫХ ПРОГРАММ")
    root.geometry("450x200") 

    message = StringVar()

    def start_parsing():
        url = entry.get()
        save_in_exel(reading(url, 'arp'), reading(url, 'rp'), reading(url, 'rpv'), reading(url, 'pp'), url)

    label = ttk.Label(text = "Введите ссылку на страницу", font=("Arial", 14))
    label.pack(anchor=N, padx=6, pady=6)
 
    entry = ttk.Entry()
    entry.pack(anchor=N, fill = X)
  
    btn = ttk.Button(text="Начать парсить", command=start_parsing)
    btn.pack(anchor=N, fill = X )

    label_mes = ttk.Label(font = ("Arial", 12), textvariable=message)
    label_mes.pack(anchor= "center")
  
    root.mainloop()
    
    
       
    
