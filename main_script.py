from docx import Document
import re
# В консолі вашої ide пропишіть pip install python-docx
# Вставте текст з лабораторної роботи в такому форматі (НЕ чіпайте дужки, просто вставляйте текст між ними, як у прикладі)
input_string = """
Куряча грудинка 65 55
Для соусу:
Часник 65 60
Базилік 50 45
Кедрові горіхи
(соняшникове насіння)
40 40
Пармезан 100 100
Олія оливкова 40 40
Маса соусу - 245
Чіабатта 200 200
"""

def is_float(value):
    try:
        float_number = float(value.replace(",", "."))
        float(float_number)
        return True
    except ValueError:
        return False

lines = input_string.strip().split('\n')

input_string = input_string.replace('\n', ' ')

pattern = r'(.+?)\s+([\d.,]+)\s+([\d.,]+)'
items = re.findall(pattern, input_string)


print(input_string)

table_data = []
for item in items:
    item_name = item[0]
    brutto = item[1]
    netto = item[2]
    if brutto.isdigit():
        brutto_multiplied = str(int(brutto) * 3)

    elif is_float(brutto):
        brutto = brutto.replace(",", ".")
        brutto_multiplied = str(round(float(brutto) * 3, 1))

    if brutto == "шт.":
        brutto = f'{item_name.split()[-1]} шт.'
        brutto_multiplied = f'{int(item_name.split()[-1])*3} шт.'
        item_name = ' '.join(item_name.split()[:-1])

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


doc.save("output.docx")
