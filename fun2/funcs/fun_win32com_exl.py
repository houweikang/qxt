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


# excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
# excel.DisplayAlerts = 0
# excel.SheetsInNewWorkbook = 1
#
#
# def screen_updating(bol=True):
#     excel.ScreenUpdating = bol
#     excel.Visible = bol

class Excel:
    def __init__(self):
        self.excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        self.excel.DisplayAlerts = 0
        self.excel.SheetsInNewWorkbook = 1

    def screen_updating(self,bol=True):
        self.excel.ScreenUpdating = bol
        self.excel.Visible = bol

    def WorkBooks(self, name_or_index):
        try:
            self.wb = self.excel.Workbooks(name_or_index)
            return self.wb
        except:
            print('Error!')

    def active_workbook(self):
        self.wb = self.excel.ActiveWorkbook
        return self.wb

    def workbooks_add(self):
        self.wb = self.excel.Workbooks.Add()
        return self.wb

    def workbooks_save(self, name, path):
        MyType = '.xlsx'
        if not name.endswith(MyType):
            name = name + MyType
        full_name = file_rename(path, name)
        self.wb.SaveAs(full_name)
        self.wb.Save()
        return self.wb

    def activesheet(self):
        self.sht = self.excel.ActiveWorkbook.ActiveSheet
        return self.sht

    def sheets_add(self, sheet_name, index=None):
        sheets_count = self.wb.Worksheets.Count
        if (not index) or (index > sheets_count):
            self.wb.Sheets.Add(After=self.wb.Sheets(sheets_count)).Name = sheet_name
        else:
            self.wb.Sheets.Add(Before=self.wb.Sheets(index)).Name = sheet_name
        return self.wb.Worksheets(sheet_name)

    def sheet_tab_color(self, sheet_name_or_index, color=7697919):
        '''
        :param sheet_name_or_index:
        :param color:7697919-红  255-深红  11854022-绿  2315831-深绿
        :return:
        '''
        color_dic = {'R': 7697919, 'DR': 255, 'G': 11854022, 'DG': 2315831}
        if isinstance(color, str):
            color = color_dic[color]
        self.wb.Sheets(sheet_name_or_index).Tab.Color = color

    def sheets_delete(self, sheet_name_or_index):
        try:
            self.wb.Sheets(sheet_name_or_index).Delete()
        except Exception as err:
            print(err)

    def sheets_select(self, sheet_name_or_index):
        try:
            self.wb.Sheets(sheet_name_or_index).Select()
        except Exception as err:
            print(err)

    def write_data(self,sht, position, data):
        if isinstance(position, str):
            cel = sht.Range(position)
        elif isinstance(position, (tuple, list)):
            cel = sht.Cells(position[0], position[1])
        try:
            if isinstance(data, (str, int, float)):
                cel.Value = data
            elif isinstance(data, (tuple, list)):
                shape = (np.array(data)).shape
                for i in range(shape[0]):
                    if len(shape) == 2:
                        for j in range(shape[1]):
                            cel.GetOffset(i, j).Value = data[i][j]
                    else:
                        cel.GetOffset(0, i).Value = data[i]
        except Exception as err:
            print('err:', err)


def file_rename(path, file_name):
    full_name = os.path.join(path, file_name)
    only_file_name, extention_name = os.path.splitext(file_name)
    i = 1
    while os.path.isfile(full_name):
        new_file_name = '{}{}{}'.format(only_file_name, '(%d)' % i, extention_name)
        full_name = os.path.join(path, new_file_name)
        i += 1
    return full_name
