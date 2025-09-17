import openpyxl
from openpyxl import Workbook, load_workbook
from bs4 import BeautifulSoup

my_file = load_workbook('excel_report.xlsx')
sheet = my_file.active


top_pg = []

for i in range(1, sheet.max_row+1):
    if i >= 4:
        break
    if i == 1:
            continue
    inner_arr = []

    for j in range(1, sheet.max_column+1):
        
        if j == 2:
            continue
        cell_obj = sheet.cell(row=i, column=j).value
        if j == 4 or j == 5:
             cell_obj_num = float(str(cell_obj).strip())
             inner_arr.append(cell_obj_num * 100)
        else:
            inner_arr.append(cell_obj)

    if inner_arr:
        top_pg.append(inner_arr)
        

print(top_pg)


with open('index.html', 'r', encoding='utf-8') as file:
    content = file.read()

soup = BeautifulSoup(content, 'html.parser')



row_ids = ['first_row', 'second_row', 'third_row']

for i in range(0, len(top_pg)):
        
        tr_tag = soup.find('tr', id=row_ids[i])
        tr_tag.clear()

        for j in range(0,4): 
             new_td = soup.new_tag('td')
             new_td.string = str(top_pg[i][j])
             tr_tag.append(new_td)


with open('index.html', 'w', encoding='utf-8') as file:
    file.write(str(soup))
