# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
from openpyxl.writer.excel import ExcelWriter
from openpyxl import load_workbook  
from openpyxl.utils import coordinate_from_string, column_index_from_string, get_column_letter

#登录部分

root_url = 'http://172.16.203.12/zentao/user-login.html'
index_url = 'http://172.16.203.12/zentao/project-task-206.html'
UA = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36"

header = {"User-Agent": UA,
           "referer":"http://172.16.203.12/zentao/my/"
           }

r = requests.Session()
f = r.get(root_url,headers = header)

r.cookies = requests.utils.cookiejar_from_dict({
    'zentaosid':'0qorkcc8602i5gg00seenj6kc2'})

r.post(root_url,
       cookies = r.cookies,
       headers = header
       )

#抓取数据部分

f = r.get(index_url,headers = header)

soup = BeautifulSoup(f.content,'lxml')

plans = soup.find_all('tr',class_='text-center')
for plan in plans:
    l = list(plan.stripped_strings)
    name = l[1]
    status = l[2]
    plan = l[-1]
    print name,status,plan,'\n'




'''
#填写表格
def style_range(ws, cell_range, style):
    """
    :param ws:  Excel worksheet instance
    :param range: An excel range to style (e.g. A1:F20)
    :param style: An openpyxl Style object
    """

    start_cell, end_cell = cell_range.split(':')
    start_coord = coordinate_from_string(start_cell)
    start_row = start_coord[1]
    start_col = column_index_from_string(start_coord[0])
    end_coord = coordinate_from_string(end_cell)
    end_row = end_coord[1]
    end_col = column_index_from_string(end_coord[0])

    for row in range(start_row, end_row + 1):
        for col_idx in range(start_col, end_col + 1):
            col = get_column_letter(col_idx)
            ws.cell('%s%s' % (col, row)).style = style

wb = load_workbook(filename = r'OE_Cloud_PMC_2017Q3.xlsx')
ws = wb.active
ws.merge_cells('B22:D22')
my_cell = ws.cell('B22')
my_cell.value = name

style_range(ws,'B22:D22',style(alignment=Alignment(horizontal='center'),
                               border=Border(top=Side(border_style='thin', color=colors.BLACK),
                                             left=Side(border_style='thin', color=colors.BLACK),
                                             bottom=Side(border_style='thin', color=colors.BLACK),
                                             right=Side(border_style='thin', color=colors.BLACK), )),)
wb.save('OE_Cloud_PMC_2017Q3.xlsx')
'''
