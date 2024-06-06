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

    message.set("��������� �������...")
    for pdf_file in files:  
        response = requests.get(pdf_file)
        pdf_content = response.content

        pdf_stream = BytesIO(pdf_content)
        document = fitz.open(stream=pdf_stream, filetype="pdf")
        try:
            #����� �������� ����������
            text = ""
            for page_num in range(len(document)):
                page = document.load_page(page_num)
                text += page.get_text()
            if ID == 'rpv':
                pattern = re.compile(r'(�������)\s+(���������)\s+(����������)', re.IGNORECASE)
                matches = re.findall(pattern, text)
                result = [' '.join(match) for match in matches]
            else:
                pattern = re.compile(r'�([A-Z�-ߨ][^0-9��]*)\s*([^0-9��]*?)�', re.IGNORECASE)
                matches = re.findall(pattern, text)
                result = [''.join(match) for match in matches]
                filtered_result = [word for word in result if word and word[0].isupper() and not word[0].isdigit()]
                result = filtered_result
                for i in range(len(result)):
                    result[i] = re.sub("\n","",result[i])

            #print(result)
            #input()
            try:
                #����� ������������ ������ 
                for res in result:
                        if res != "����������� ��������������� ����������� �����������":
                            name_list.append(res)
                            break
            except:
                name_list.append(f"�� ������� ��������� ����: {pdf_file}")
                print(f"�� ������� ��������� ����: {pdf_file}")
        finally:
            document.close()  
    return name_list

def find_pdf(url, ID):
    error = False

    print("������ ��������...")
    message.set("������ ��������...")
    # id ��� ��������� "arp", ������� �������� - "rp", ��� ���� ���������� - "rpv", ��������� ������� - "pp"
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
    
    param = "���-��"
    count = 0
    
    if ID == 'arp':
        param = "���������"
        global count_annot
        global a_link_list
        count_annot = len(link_to_pdf)
        a_link_list = link_to_pdf
        count = count_annot
    elif ID == 'rp':
        param = "������� ��������"
        global count_rab
        global r_link_list
        count_rab = len(link_to_pdf)
        r_link_list = link_to_pdf
        count = count_rab
    elif ID == 'rpv':
        param = "�������� ����������"
        global count_rpv
        global v_link_list
        count_rpv = len(link_to_pdf)
        v_link_list = link_to_pdf
        count = count_rpv
    elif ID == 'pp':
        param = "�������� �������"
        global count_pp
        global p_link_list
        count_pp = len(link_to_pdf)
        p_link_list = link_to_pdf
        count = count_pp
    
    if (error):
        print("��������, ����� ������ ��� 0_0")
    else:
        print(f"���� ������� {param}: {count}")

    message.set(f"���� ������� {param}: {count}")

    return link_to_pdf

def save_in_exel(list_annot, list_rab, list_rpv, list_pp, url):

    global count_annot, count_rab, count_pp, count_rpv
    global a_link_list, r_link_list, v_link_list, p_link_list

    count_annot_string = f"{count_annot}"
    count_rab_string = f"{count_rab}"
    count_rpv_string = f"{count_rpv}"
    count_pp_string = f"{count_pp}"

    df1 = pd.DataFrame(list_annot, columns=['��������� ' + "(" + str(count_annot) + ")"])
    df1.insert(1, "������", a_link_list)
    df2 = pd.DataFrame(list_rab, columns = ['������� ��������� ' + "(" + str(count_rab) + ")"])
    df2.insert(1, "������", r_link_list)
    df3 = pd.DataFrame(list_rpv, columns=['��������� ���������� ' + "(" + str(count_rpv) + ")"])
    df3.insert(1, "������", v_link_list)
    df4 = pd.DataFrame(list_pp, columns=['��������� ������� ' + "(" + str(count_pp) + ")"])
    df4.insert(1, "������", p_link_list)

    response = requests.get(url)
    soup = BeautifulSoup(response.text, "lxml")
    title = soup.find('a', class_ = 'active', string = re.compile('�����������'))

    filename = f'{title.text}.xlsx'
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='���������', index=False)
        df2.to_excel(writer, sheet_name='������� ���������', index=False)
        df3.to_excel(writer, sheet_name='��������� ����������', index=False)
        df4.to_excel(writer, sheet_name='��������� �������', index=False)

    print(f"������ ����������� ������� �������� � '{filename}'")

    filepath = os.path.abspath(filename)
    adjust_column_width(filepath, '���������')
    adjust_column_width(filepath, '������� ���������')
    adjust_column_width(filepath, '��������� ����������')
    adjust_column_width(filepath, '��������� �������')
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
    root.title("������ ��������������� ��������")
    root.geometry("450x200") 

    message = StringVar()

    def start_parsing():
        url = entry.get()
        save_in_exel(reading(url, 'arp'), reading(url, 'rp'), reading(url, 'rpv'), reading(url, 'pp'), url)

    label = ttk.Label(text = "������� ������ �� ��������", font=("Arial", 14))
    label.pack(anchor=N, padx=6, pady=6)
 
    entry = ttk.Entry()
    entry.pack(anchor=N, fill = X)
  
    btn = ttk.Button(text="������ �������", command=start_parsing)
    btn.pack(anchor=N, fill = X )

    label_mes = ttk.Label(font = ("Arial", 12), textvariable=message)
    label_mes.pack(anchor= "center")
  
    root.mainloop()
    
    
       
    
