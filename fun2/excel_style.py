import win32com.client
from win32com.client import constants as c  # 旨在直接使用VBA常数
import os

excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
excel.Visible = 1
excel.DisplayAlerts = 0
excel.ScreenUpdating = 1


def style(rg, amount=1, title_rows=2):
    '''
    两个主要格式 1：数据量较大 0：较少 交叉 标题，日期【副标题】
    '''
    rg = rg
    rs = rg.Rows.Count
    cs = rg.Columns.Count
    font_name = 'Microsoft YaHei UI'

    rg.Borders(c.xlEdgeBottom).LineStyle = c.xlNone
    rg.HorizontalAlignment = c.xlCenter
    rg.VerticalAlignment = c.xlCenter

    if title_rows == 2:
        tl_rg = rg.Rows(1)
        dt_rg = rg.Rows(2)
        cl_rg = rg.Rows(3)
        data_rg = rg.Range(rg.Cells(4, 1), rg.Cells(rs, cs))
    elif title_rows == 1:
        tl_rg = rg.Rows(1)
        cl_rg = rg.Rows(2)
        data_rg = rg.Range(rg.Cells(3, 1), rg.Cells(rs, cs))
    elif title_rows == 0:
        cl_rg = rg.Rows(1)
        data_rg = rg.Range(rg.Cells(2, 1), rg.Cells(rs, cs))

    if tl_rg:
        if tl_rg.Cells(1).Count == 1:
            tl_rg.Merge()
        font = tl_rg.Font
        font.Name = font_name
        font.Size = 16
        font.Bold = True
        font.Color = 7884319
        tl_rg.Rows.RowHeight = 30

    if dt_rg:
        dt_rg.Cells(1, cs).Value = dt_rg.Cells(1, 2).Value
        rg.Range(dt_rg.Cells(1), dt_rg.Cells(cs - 1)).Merge()
        dt_rg.Cells(1).HorizontalAlignment = c.xlRight
        dt_rg.Cells(cs).HorizontalAlignment = c.xlLeft
        font = dt_rg.Font
        font.Size = 12
        font.Name = font_name

    if cl_rg:
        font = cl_rg.Font
        font.Name = font_name
        font.Bold = True
        font.Color = 16777215
        cl_rg.Interior.Color = 7884319

        if amount == 1:  # 数据量大
            font.Size = 10
            cl_rg.Cells.EntireRow.AutoFit()
            cl_rg.Cells.EntireColumn.AutoFit()
        elif amount == 0:  # 数据量小
            font.Size = 12
            cl_rg.Rows.RowHeight = 30

    if data_rg:
        font = data_rg.Font
        font.Name = font_name
        data_rg.Borders.LineStyle = c.xlContinuous
        data_rg.Borders.Weight = c.xlThin
        data_rg.Borders(c.xlInsideHorizontal).Color = 13224393
        data_rg.Borders(c.xlInsideVertical).Color = 13224393

        if amount == 1:  # amount 1 数据量大;
            font.Size = 10
            data_rg.Cells.EntireColumn.AutoFit()
            data_rg.Cells.EntireRow.AutoFit()
        elif amount == 0:  # 数据量少；
            font.Size = 11
            data_rg.Columns.ColumnWidth = 15
            data_rg.Rows.RowHeight = 30
            data_rg.Borders.ThemeColor = 1
            data_rg.Borders.TintAndShade = -0.14996795556505
            data_rg.Borders(c.xlEdgeBottom).Color = 5
            data_rg.Borders(c.xlEdgeBottom).TintAndShade = -0.499984740745262


def autofit(rng, columnlist):
    col_inds = []
    # 将自动调整列宽的列下标写进列表
    if isinstance(columnlist, list):
        col_inds = columnlist
    elif isinstance(columnlist, int):
        col_inds.append(columnlist)
    for col_ind in col_inds:
        rng.Columns(col_ind).AutoFit()


def bold(rng, tag, list):
    rc = rng.Rows.Count
    for i in range(1, rc + 1):
        for j in list:
            if rng.Cells(i, j).Value:
                if rng.Cells(i, j).Value.find(tag) >= 0:
                    rng.Rows(i).Font.Bold = True
                    continue
    #
    # def merge(self, rng, column_list):
    #     rc = rng.Rows.Count
    #     for j in column_list:
    #         for i in range(rc, 1, -1):
    #             if j == 1:
    #                 if not (rng.Cells(i, j).Value and rng.Cells(i - 1, j).Value):
    #                     # rng.Cells(i-1,j).Select()
    #                     self.excelrng.Cells(i, j), rng.Cells(i - 1, j)).Merge()
    #                     elif rng.Cells(i, j).Value == rng.Cells(i - 1, j).Value:
    #                     self.sheet.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()
    #                     elif j > 1:
    #                     if self.sheet.Range(rng.Cells(i, j - 1), rng.Cells(i - 1, j - 1)).MergeCells:
    #                         if
    #                     (not rng.Cells(i, j).Value) and (not rng.Cells(i - 1, j).Value):
    #                     self.sheet.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()
    #                     elif rng.Cells(i, j).Value == rng.Cells(i - 1, j).Value: \
    #                         self.sheet.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()
    #
    # def xl3Triangles(self, rng, columnlist):
    #     col_inds = []
    #     if isinstance(columnlist, list):
    #         col_inds = columnlist
    #     elif isinstance(columnlist, int):
    #         col_inds.append(columnlist)
    #     for col_ind in col_inds:
    #         rng_col = rng.Columns(col_ind)
    #         rng_col.FormatConditions.AddIconSetCondition()
    #         rng_col_fc1 = rng.FormatConditions(1)
    #         rng_col_fc1.IconSet = self.workbook.IconSets(c.xl3Triangles)
    #         rng_col_fc1_ic2 = rng_col_fc1.IconCriteria(2)
    #         rng_col_fc1_ic3 = rng_col_fc1.IconCriteria(3)
    #         rng_col_fc1_ic2.Type = c.xlConditionValueNumber
    #         rng_col_fc1_ic2.Operator = 7
    #         rng_col_fc1_ic2.Value = 0
    #         rng_col_fc1_ic3.Type = c.xlConditionValueNumber
    #         rng_col_fc1_ic3.Operator = 5
    #         rng_col_fc1_ic3.Value = 0
    #
    # def xlConditionValueNumber(self, rng, columnlist):
    #     col_inds = []
    #     if isinstance(columnlist, list):
    #         col_inds = columnlist
    #     elif isinstance(columnlist, int):
    #         col_inds.append(columnlist)
    #     for col_ind in col_inds:
    #         rng_col = rng.Columns(col_ind)
    #         rng_col.Style = "Percent"
    #         rng_col.FormatConditions.AddDatabar()
    #         rng_col_fc1 = rng.FormatConditions(1)
    #         rng_col_fc1.MinPoint.Modify(newtype=c.xlConditionValueNumber, newvalue=0)
    #         rng_col_fc1.MaxPoint.Modify(newtype=c.xlConditionValueNumber, newvalue=1)
    #         rng_col_fc1.BarColor.Color = 13012579
    #         rng_col_fc1.BarColor.TintAndShade = 0
    #
    # def none_gridlines(self):
    #     self.excel.ActiveWindow.DisplayGridlines = False


def add_title_and_subtitle(wb_path, sheet_name_or_index, title, subtitle, label='统计时间：'):
    wb = excel.Workbooks.Open(wb_path)
    sht = wb.Sheets(sheet_name_or_index)
    sht.Cells(1, 1).Value = title
    sht.Cells(2, 1).Value = label
    sht.Cells(2, 2).Value = subtitle
    wb.Save()


def hour_style(wb_path, sheet_name_or_index):
    # 无网格线
    wb = excel.Workbooks.Open(wb_path)
    sht = wb.Sheets(sheet_name_or_index)
    excel.ActiveWindow.DisplayGridlines = False
    style(sht.UsedRange, amount=0)
    # autofit(sht.UsedRange,)


if __name__ == '__main__':
    # rg=ExcelStyle()
    # rg.hour_style
    path = r'e:\报表\晨报小时报模板\use'
    wb = new_workbook('a.xlsx', path)
