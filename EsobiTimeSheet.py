#!/usr/local/bin/python3.7
# -*- coding: utf-8 -*-
import pandas as pd
import pathlib
from sys import platform
from os import listdir
from os.path import isfile, join

path = str(pathlib.Path().absolute())
prefix = '/'
if platform == 'darwin':
    path = '/Users/belindalu/Documents/EsobiTimeSheet'
if platform == 'win32':
    prefix = '\\'
path = path + prefix


def execute():
    # 1. 取得time_sheet_path
    time_sheet_path = foo()

    # 2. 取得檔案內容
    time_sheet = pd.read_excel(time_sheet_path)
    export_sheet = pd.DataFrame()
    export_sheet['member'], export_sheet['work_time'] = time_sheet['登記人'], time_sheet['耗時']
    export_sheet = export_sheet.groupby(['member']).sum()

    # 3. 取得member_path

    member_path = path + 'members.xlsx'

    # 4. 篩選出今日上班人員
    members = pd.read_excel(member_path)
    members = members[members.check == 'Y'].drop('check', axis=1)

    # 5. 排除重複輸入的人員
    members = members.drop_duplicates(keep='first')

    # 6. join 人員清單與工時
    final_data = pd.merge(members, export_sheet, how='left', on='member')
    final_data.fillna(0, inplace=True)
    final_data = final_data[final_data.work_time < 8]
    final_data['loss_time'] = 8 - final_data['work_time']
    print(final_data)


def foo():
    files = [f for f in listdir(path) if isfile(join(path, f))]
    file_dict = {}
    i = 0
    for f in files:
        file_dict[i] = f
        print(i, f)
        i = i + 1
    file_id = int(input('請輸入time_sheet檔對應的id' + '\n' + '檔案id:').strip())
    return path + file_dict[file_id]


try:
    execute()
except:
    print('你他媽的別在那邊亂喔...')
