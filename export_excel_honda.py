#!/usr/bin/env python
# -*- coding: utf-8 -*-
import pandas as pd
import datetime as dt
import os
# all_file = "~/Desktop/price_analyse/export_excel_data/all_file_name.xlsx"
# df = pd.read_excel(all_file)
# print(df2)
# f_list = []
# for i in df2:
#     f_list.append(i)
is_honda = True
month = 'may'
year = '2021'
extend = '_honda' if is_honda else ''
file_extention = '.xlsx' if is_honda else '.xls'
source_name = f'{month}{year}{extend}'
result_file_name = f'sum_{source_name}.csv'
dir_path = os.path.dirname(os.path.realpath(__file__))
path = f'{dir_path}/{source_name}/'
f_list = os.listdir(f'{path}')
print(f'{len(f_list)}')

df = []
date = []
pv_no = []
tax_id = []
branch = []
cus_name = []
amount = []
vat = []
total = []
for i in f_list:
    df = pd.read_excel(path+i,sheet_name=2)
    date.append(str(df.loc[11,'Unnamed: 22']))
    pv_no.append(str(df.loc[10,'Unnamed: 22']))
    tax_id.append('0' + str(df.loc[13,'Unnamed: 6']))
    if df.loc[13,'Unnamed: 10'] == 'x' or df.loc[13,'Unnamed: 10'] == 'X':
        branch.append('สำนักงานใหญ่')
    elif df.loc[13,'Unnamed: 13'] == "x" or df.loc[13,'Unnamed: 13'] == "X":
        branch.append(str(df.loc[13,'Unnamed: 16']))
    cus_name.append(str(df.loc[10,'Unnamed: 5']))
    amount.append(str(df.loc[38,'Unnamed: 21']))
    vat.append(str(df.loc[39,'Unnamed: 21']))
    total.append(str(df.loc[40,'Unnamed: 21']))

print(len(date))
print(len(pv_no))
print(len(tax_id))
print(len(branch))
print(len(cus_name))
print(len(amount))
print(len(vat))
print(len(total))

dict1 = {
    'date': date,
    'pv_no': pv_no,
    'tax_id': tax_id,
    'branch': branch,
    'cus_name': cus_name,
    'amount': amount,
    'vat': vat,
    'total': total
    }

df = pd.DataFrame(dict1, columns=['date','pv_no','tax_id','branch','cus_name','amount','vat','total'])
df.to_csv(result_file_name, encoding='utf-8-sig')
