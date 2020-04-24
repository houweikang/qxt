#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2020/4/24 9:33:13
# @Author  : HouWk
# @Site    : 
# @File    : fun_win32com_exl_chart.py
# @Software: PyCharm
import win32com.client
from win32com.client import constants as c  # 旨在直接使用VBA常数
import os
import numpy as np

excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
excel.Visible = 1
excel.DisplayAlerts = 0
excel.ScreenUpdating = 1
excel.SheetsInNewWorkbook = 1


def WorkBooks(name_or_index):
    try:
        return excel.Workbooks(name_or_index)
    except:
        print('Error!')


def active_workbook():
    return excel.ActiveWorkbook


def workbooks_add(sheet1_name=None):
    if sheet1_name:
        wb = excel.Workbooks.Add()
        wb.Sheets(1).Name = sheet1_name
        return wb.Sheets(sheet1_name)
    else:
        wb = excel.Workbooks.Add()
        return wb


def workbooks_save(wb, name, path):
    wb = wb
    MyType = '.xlsx'
    if not name.endswith(MyType):
        name = name + MyType
    full_name = os.path.join(path, name)
    wb.SaveAs(full_name)
    wb.Save()
    return wb


def activesheet():
    return excel.ActiveWorkbook.ActiveSheet


def sheets_add(wb, sheet_name, index=None):
    sheets_count = wb.Worksheets.Count
    if (not index) or (index > sheets_count):
        wb.Sheets.Add(After=wb.Sheets(index)).Name = sheet_name
    else:
        wb.Sheets.Add(Before=wb.Sheets(index)).Name = sheet_name
    return wb.Worksheets(sheet_name)


def write_data(sht, position, data):
    if isinstance(position, str):
        cel = sht.Range(position)
    elif isinstance(position, (tuple, list)):
        cel = sht.Cells(position[0], position[1])
    if cel:
        if isinstance(data, (str, int, float)):
            cel.Value = data
        elif isinstance(data, (tuple, list)):
            shape = (np.array(data)).shape
            for i in range(shape[0]):
                if len(shape)==2:
                    for j in range(shape[1]):
                        cel.GetOffset(i, j).Value = data[i][j]
                else:
                    cel.GetOffset(0, i).Value = data[i]
    else:
        print('{}有问题！'.format(position))


def main():
    sht = activesheet()
    write_data(sht, 'a1', '1')
    write_data(sht, [2, 2], [[1, 2, 3], [4, 5, 6]])
    # print(sht.Name)


if __name__ == "__main__":
    main()
