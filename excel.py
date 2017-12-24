# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import xlrd
import re
from openpyxl import Workbook
origin = 'origin.xlsx'
fd = xlrd.open_workbook(origin)
wb = Workbook()
ws = wb.active
key_dict = {}

for sheet in fd.sheets():
    if sheet.nrows < 2:
        continue
    row_to_write = ws.max_row + 1
    for index, item in enumerate(sheet.col_values(0)):
        if item == '':
            continue
        if isinstance(item, int) or isinstance(item, float):
            key = index
            val = item
        else:
            k, *v = re.split(': |： |:|：', item)
            if v == []:
                if index < 10:
                    key = index
                    val = k
                else:
                    key = k
                    val = ''
            else:
                key = k
                val = v[0]
        col_n = key_dict.get(key, None)
        if col_n == None:
            col_n = len(key_dict) + 1
            key_dict[key] = col_n
        ws.cell(row = row_to_write, column = col_n).value = val
for k, v in key_dict.items():
    ws.cell(row = 1, column = v).value=k
wb.save(filename = 'test.xlsx')
            
    
    
