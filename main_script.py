from docx import Document
# в консолі вашої ide пропишіть pip install python-docx
# вставте текст з лабораторної роботи в такому форматі (НЕ чіпайте дужки, просто вставляйте текст між ними, як у прикладі)
input_string = """
Філе куряче 80 70
Мікс салат 70 65
Помідори черрі 35 30
Сир Пармезан 30 30
Бекон 43 40
Батон 35 30
Часник сушений 2 2
Яйця перепелині 2 шт. 35
Для соусу
Майонез 12 12
Лимонний сік 4 4
Часник 3 2
Сир «Пармезан» 22 20
Олія оливкова 12 12
"""


lines = input_string.split('\n')


table_data = []
for line in lines:
    items = line.split()
    if len(items) == 3:
        item_name = ' '.join(items[:-2])
        brutto = items[-2] if items[-2].isdigit() else items[-2] + " шт."
        netto = items[-1]
        brutto_multiplied = str(int(brutto) * 3) if brutto.isdigit() else brutto
        netto_multiplied = int(netto) * 3
        table_data.append([item_name, brutto, netto, brutto_multiplied, netto_multiplied])
    elif len(items) == 4:
        item_name = ' '.join(items[:-3])
        count = items[-3]
        brutto = items[-2]
        netto = items[-1]
        brutto_multiplied = int(brutto) * 3
        netto_multiplied = int(netto) * 3
        table_data.append([f"{item_name} {count}", brutto, netto, brutto_multiplied, netto_multiplied])

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
