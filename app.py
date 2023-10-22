import os
from docx import Document
import tkinter as tk
from tkinter import messagebox
from tkinter import Label
import re


def generate_docx():
    try:
        input_string = input_text.get("1.0", "end-1c")

        def is_float(value):
            try:
                float_number = float(value.replace(",", "."))
                float(float_number)
                return True
            except ValueError:
                return False


        # Iterate through lines and merge lines that do not start with a number with the previous line
        new_string = input_string.replace('\n', ' ')

        pattern = r'(.+?)\s+((?:\d+(?:[.,]\d+)?)\s*(?:шт\.?)?)\s+([\d.,]+)'
        items = re.findall(pattern, new_string)

        print(new_string)
        print(items)

        table_data = []
        for item in items:
            item_name = item[0]
            brutto = item[1]
            print(brutto)
            netto = item[2]
            if brutto.isdigit():
                brutto_multiplied = str(int(brutto) * 3)

            elif is_float(brutto):
                brutto = brutto.replace(",", ".")
                brutto_multiplied = str(round(float(brutto) * 3, 1))

            if "шт" in brutto:
                if brutto.split(' ')[0].isdigit():
                    brutto_multiplied = f'{str(int(brutto.split(" ")[0]) * 3)} шт.'

                elif is_float(brutto.split(' ')[0]):
                    brutto_multiplied = f'{str(round(float(brutto.split(" ")[0]) * 3, 1))} шт.'

            if netto.isdigit():
                netto_multiplied = str(int(netto) * 3)

            elif is_float(netto):
                netto = netto.replace(",", ".")
                netto_multiplied = str(round(float(netto) * 3, 1))

            table_data.append([item_name, brutto, netto, brutto_multiplied, netto_multiplied])

        doc = Document()
        table = doc.add_table(rows=1, cols=6)

        table_header = table.rows[0].cells
        table_header[0].text = 'Назва'
        table_header[1].text = 'Одиниця виміру'
        table_header[2].text = 'Брутто'
        table_header[3].text = 'Нетто'
        table_header[4].text = 'Брутто * 3'
        table_header[5].text = 'Нетто * 3'

        for row_data in table_data:
            row_cells = table.add_row().cells
            row_cells[0].text = row_data[0]
            row_cells[1].text = 'г'
            row_cells[2].text = row_data[1]
            row_cells[3].text = row_data[2]
            row_cells[4].text = str(row_data[3])
            row_cells[5].text = str(row_data[4])

        # Get the desktop path
        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

        # Find the next available filename for the table
        i = 1
        while True:
            filename = f"table_{i}.docx"
            filepath = os.path.join(desktop_path, filename)
            if not os.path.exists(filepath):
                break
            i += 1

        # Save the document
        doc.save(filepath)
        messagebox.showinfo("Info", f"Таблиця '{filename}' згенерована")
    except UnboundLocalError:
        messagebox.showerror("Error", f'Якщо в графі для "Брутто" є дані, що включають "шт", перевірте, чи є між числом та "шт" пробіл')

# GUI Setup
root = tk.Tk()
root.title("Table autofiller")
initial_width = 600  # Set your desired initial width
initial_height = 500  # Set your desired initial height
root.geometry(f"{initial_width}x{initial_height}")

input_label = tk.Label(root, text="""Вставте скопійовані з лабораторної роботи дані в форматі
Назва_сировини брутто нетто
Назва_сировини брутто нетто 
і так далі...

Бажано видалити рядки типу "Для соусу" 
та тп, що не мають формату Сирована - Маса - Маса)

Ніяк форматувати текст самостійно не треба, 
(наприклад: назва на двох рядках, або маси на новому рядку після сировини) 
в скріпті вже прописане таке форматування

↓↓↓↓↓
""")
input_label.pack()

input_text = tk.Text(root, height=10, width=50)
input_text.pack()

output_label = Label(root, text="Таблиця зберігається на робочий стіл")
output_label.pack()

generate_button = tk.Button(root, text="Створити .docx з таблицею", command=generate_docx)
generate_button.pack()


root.mainloop()