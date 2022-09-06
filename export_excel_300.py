#!/usr/bin/env python
# -*- coding: utf-8 -*-
import pandas as pd
import datetime as dt
import os


month = 'aug'
year = '22'
is_honda_list = [True, False]
for honda in is_honda_list:
    is_honda = honda
    print(f'is_honda = {is_honda}')
    extend = '_honda' if is_honda else ''
    file_extention = '.xlsx' if is_honda else '.xls'
    file_name = f'{month}{year}{extend}'
    source_name = f'raw_data/{file_name}'
    result_file_name = f'result/sum_{file_name}.csv'
    dir_path = os.path.dirname(os.path.realpath(__file__))
    path = f'{dir_path}/{source_name}/'
    f_list = os.listdir(f'{path}')
    f_list.sort()
    print(type(f_list))
    for i in f_list:
        if not i.lower().endswith((file_extention)):
            print(f'removing {i}')
            f_list.remove(i)
    print(f'{len(f_list)}')
    print(f'{path}')

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
        print(f'{path+i}')
        if is_honda:
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
        else:
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
