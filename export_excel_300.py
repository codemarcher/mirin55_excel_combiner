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
# path = "~/Desktop/price_analyse/export_excel_data/dec2020/"
result_file_name = 'sum_feb2021.csv'
path = "feb2021/"
f_list = []
for filename in os.listdir(path):
    if filename.endswith(".xls"):
        f_list.append(filename)
        print(f'filename: {filename}')
    else:
        continue

# f_list = []
# for i in range(len(df)):
#     f_list.append(df.loc[i, 'all_file'])
# print(f_list)
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
    date.append(str(pd.to_datetime(df.loc[12,'Unnamed: 22']).strftime("%d/%m/%Y")))
    pv_no.append(str(df.loc[11,'Unnamed: 22']))
    tax_id.append('0' + str(df.loc[15,'Unnamed: 6']))
    if df.loc[15,'Unnamed: 10'] == 'x' or df.loc[15,'Unnamed: 10'] == 'X':
        branch.append('สำนักงานใหญ่')
    elif df.loc[15,'Unnamed: 13'] == "x" or df.loc[15,'Unnamed: 13'] == "X":
        branch.append(str(df.loc[15,'Unnamed: 16']))
    cus_name.append(str(df.loc[11,'Unnamed: 5']))
    amount.append(str(df.loc[40,'Unnamed: 21']))
    vat.append(str(df.loc[40,'Unnamed: 21']))
    total.append(str(df.loc[40,'Unnamed: 21']))

print(len(date))
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

# path = "~/Desktop/price_analyse/export_excel_data/09.September2019/"

# file_1 = "MRHD2019-09-01 - Honda Suwinthawong.xls"
# file_2 = "MRHD2019-09-02 - Honda Ladkrabang.xls"
# file_3 = "MRHD2019-09-03 - Honda Klongluang.xls"
# file_4 = "MRHD2019-09-04 - Honda Thanyaburi.xls"
# file_5 = "MRHD2019-09-05 - Honda Lumlukka.xls"
# file_6 = "MRHD2019-09-06 - Honda Suwinthawong.xls"
# file_7 = "MRHD2019-09-07 - Honda Ladkrabang.xls"
# file_8 = "MRHD2019-09-08 - Honda Klongluang.xls"
# file_9 = "MRHD2019-09-09 - Honda Thanyaburi.xls"
# file_10 = "MRHD2019-09-10 - Honda Lumlukka.xls"
# file_11 = "MRHD2019-09-11 - Honda Suwinthawong.xls"
# file_12 = "MRHD2019-09-12 - Honda Ladkrabang.xls"
# file_13 = "MRHD2019-09-13 - Honda Klongluang.xls"
# file_14 = "MRHD2019-09-14 - Honda Thanyaburi.xls"
# file_15 = "MRHD2019-09-15 - Honda Lumlukka.xls"
# file_16 = "MRHD2019-09-16 - Honda Suwinthawong.xls"
# file_17 = "MRHD2019-09-17 - Honda Ladkrabang.xls"
# file_18 = "MRHD2019-09-18 - Honda Klongluang.xls"
# file_19 = "MRHD2019-09-19 - Honda Thanyaburi.xls"
# file_20 = "MRHD2019-09-20 - Honda Lumlukka.xls"

# file_list = [file_1, file_2, file_3, file_4, file_5, file_6, file_7, file_8, file_9, file_10,
#     file_11, file_12, file_13, file_14, file_15, file_16, file_17, file_18, file_19, file_20]

# df = []
# date = []
# pv_no = []
# tax_id = []
# branch = []
# cus_name = []
# amount = []
# vat = []
# total = []
# for i in file_list:
#     df = pd.read_excel(path+i,sheet_name=2)
#     date.append(str(df.loc[12,'Unnamed: 22'].strftime("%d/%m/%Y")))
#     pv_no.append(str(df.loc[11,'Unnamed: 22']))
#     tax_id.append('0' + str(df.loc[15,'Unnamed: 6']))
#     if df.loc[15,'Unnamed: 10'] == 'x' or df.loc[15,'Unnamed: 10'] == 'X':
#         branch.append('สำนักงานใหญ๋')
#     elif df.loc[15,'Unnamed: 13'] == "x" or df.loc[15,'Unnamed: 13'] == "X":
#         branch.append(str(df.loc[15,'Unnamed: 16']))
#     cus_name.append(str(df.loc[11,'Unnamed: 5']))
#     amount.append(str(df.loc[40,'Unnamed: 21']))
#     vat.append(str(df.loc[40,'Unnamed: 21']))
#     total.append(str(df.loc[40,'Unnamed: 21']))

# # print(len(date))
# # print(len(pv_no))
# # print(len(tax_id))
# # print(len(branch))
# # print(len(cus_name))
# # print(len(amount))
# # print(len(vat))
# # print(len(total))

# dict1 = {
#     'date': date,
#     'pv_no': pv_no,
#     'tax_id': tax_id,
#     'branch': branch,
#     'cus_name': cus_name,
#     'amount': amount,
#     'vat': vat,
#     'total': total
#     }
# # print(dict1)
# df = pd.DataFrame(dict1, columns=['date','pv_no','tax_id','branch','cus_name','amount','vat','total'])
# df.to_csv('test_file3.csv', encoding='utf-8-sig')

# print(cus_name)
# df = pd.read_excel(path+file_5,sheet_name=2)
# pd.set_option('display.max_columns', df.shape[1]+1)

# cus_name = str(df.loc[11,'Unnamed: 5'])
# tax_id = str(df.loc[15,'Unnamed: 6'])
# tax_id = '0' + tax_id
# pv_no = str(df.loc[11,'Unnamed: 22'])
# date = df.loc[12,'Unnamed: 22']
# date = str(date.strftime("%d/%m/%Y"))
# total = str(df.loc[40,'Unnamed: 21'])
# vat = str(df.loc[40,'Unnamed: 21'])
# net = str(df.loc[40,'Unnamed: 21'])

# # print(cus_name + tax_id + pv_no + date + total + vat + net)

# print(dict1)
# df.to_csv('test_file5.csv', encoding='utf-8-sig')
# if df.loc[15,'Unnamed: 10'] == 'x':
#     branch.append('สำนักงานใหญ๋')
# elif df.loc[15,'Unnamed: 13'] == "x":
#     branch.append(str(df.loc[15,'Unnamed: 16']))
