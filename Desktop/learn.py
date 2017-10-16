# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
from openpyxl.writer.excel import ExcelWriter
from openpyxl import load_workbook  
from openpyxl.utils import coordinate_from_string, column_index_from_string, get_column_letter
from openpyxl.styles import Border,Side
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
    time = l[3]
    plan = l[-1]
    if plan != '100%':
        print name,status,time,plan,'\n'





#填写表格
border = Border(
    left=Side(style='medium',color='FF000000'),
    right=Side(style='medium',color='FF000000'),
    top=Side(style='medium',color='FF000000'),
    bottom=Side(style='medium',color='FF000000'),
    diagonal=Side(style='medium',color='FF000000'),
    diagonal_direction=0,outline=Side(style='medium',
    color='FF000000'),
    vertical=Side(style='medium',color='FF000000'),
    horizontal=Side(style='medium',color='FF000000'))


wb = load_workbook(filename = r'OE_Cloud_PMC_2017Q3.xlsx')
ws = wb.active
ws['B22'] = name
ws['E22'] = status + plan

wb.save('OE_Cloud_PMC_2017Q3.xlsx')


