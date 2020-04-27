import win32com.client
from win32com.client import constants as c  # 旨在直接使用VBA常数
import os

excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
excel.DisplayAlerts = 0



class RangeStyle:
    '''
    Excel区域格式
    两个主要格式 1：数据量较大 2：较少 交叉 标题，日期【副标题】
    次要格式：加粗  合并  两个条件格式
    '''

    def __init__(self, sht=None, range=None):
        if sht:
            self.sheet = sht
        else:
            self.sheet = excel.ActiveSheet
        if range:
            self.range = range
        else:
            self.range = self.sheet.UsedRange

    def none_gridlines(self):
        excel.ActiveWindow.DisplayGridlines = False

    def fontname(self, name='Microsoft YaHei UI'):  # 字体名称
        self.range.Font.Name = name

    def alignment(self):  # 对齐方式
        self.range.HorizontalAlignment = c.xlCenter
        self.range.VerticalAlignment = c.xlCenter

    def title_style(self):
        rg = self.range.Rows(1)
        if rg.Cells(1).Count == 1:
            rg.Merge()
        font = rg.Font
        font.Name = 'Microsoft YaHei UI'
        font.Size = 16
        font.Bold = True
        font.Color = 7884319
        rg.Borders(c.xlEdgeBottom).LineStyle = c.xlNone
        rg.Rows.RowHeight = 30

    def subtitle_style(self):
        rg = self.range.Rows(2)
        cols = rg.Cells.Count
        rg.Cells(1, cols).Value = rg.Cells(1, 2).Value
        self.sheet.Range(rg.Cells(1, 1), rg.Cells(1, cols - 1)).Merge()
        rg.Cells(1).HorizontalAlignment = c.xlRight
        rg.Cells(cols).HorizontalAlignment = c.xlLeft
        font = rg.Font
        font.Size = 10

    def cols_style(self, rng=None, amount=True):
        if rng:
            rg = rng
        else:
            rg = self.range.Rows(1)
        if rg.Rows.Count == 1:
            font = rg.Font
            font.Bold = True
            font.Color = 16777215
            rg.Interior.Color = 7884319

            if amount:  # 数据量大
                font.Size = 13
                rg.Rows.RowHeight = 30
            else:  # 数据量小
                font.Size = 10
                rg.Cells.EntireRow.AutoFit()
                rg.Cells.EntireColumn.AutoFit()

    def data_style(self, rng=None, amount=True, interiorcolor=False):
        if rng:
            rg = rng
        else:
            rg = self.range

        rg.Borders.LineStyle = c.xlContinuous
        rg.Borders.Weight = c.xlThin
        rg.Borders(c.xlInsideHorizontal).Color = 13224393
        rg.Borders(c.xlInsideVertical).Color = 13224393
        font = rg.Font
        if amount:  # amount 1 数据量大;
            font.Size = 11
            rg.Columns.ColumnWidth = 15
            rg.Rows.RowHeight = 30
            rg.Borders.ThemeColor = 1
            rg.Borders.TintAndShade = -0.14996795556505
            rg.Borders(c.xlEdgeBottom).Color = 5
            rg.Borders(c.xlEdgeBottom).TintAndShade = -0.499984740745262
        else:  # amount 0 数据量少 ；
            font.Size = 10
            rg.Cells.EntireColumn.AutoFit()
            rg.Cells.EntireRow.AutoFit()

        if interiorcolor:
            rc = rg.Rows.Count
            if rc > 2:
                for i in range(1, rc, 2):
                    rg.Rows(i).Interior.Color = 15921906

    def autofit(self, rng, columnlist):
        col_inds = []
        # 将自动调整列宽的列下标写进列表
        if isinstance(columnlist, list):
            col_inds = columnlist
        elif isinstance(columnlist, int):
            col_inds.append(columnlist)
        for col_ind in col_inds:
            rng.Columns(col_ind).AutoFit()

    def bold(self, rng, tag, column_list):
        rc = rng.Rows.Count
        for i in range(1, rc + 1):
            for j in column_list:
                if rng.Cells(i, j).Value:
                    if rng.Cells(i, j).Value.find(tag) >= 0:
                        rng.Rows(i).Font.Bold = True
                        continue

    def merge(self, rng, column_list):
        rc = rng.Rows.Count
        for j in column_list:
            for i in range(rc, 1, -1):
                if j == 1:
                    if not (rng.Cells(i, j).Value and rng.Cells(i - 1, j).Value):
                        # rng.Cells(i-1,j).Select()
                        self.sheet.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()
                    elif rng.Cells(i, j).Value == rng.Cells(i - 1, j).Value:
                        self.sheet.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()
                elif j > 1:
                    if self.sheet.Range(rng.Cells(i, j - 1), rng.Cells(i - 1, j - 1)).MergeCells:
                        if (not rng.Cells(i, j).Value) and (not rng.Cells(i - 1, j).Value):
                            self.sheet.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()
                        elif rng.Cells(i, j).Value == rng.Cells(i - 1, j).Value:
                            self.sheet.Range(rng.Cells(i, j), rng.Cells(i - 1, j)).Merge()

    def xl3Triangles(self, rng, columnlist):
        col_inds = []
        if isinstance(columnlist, list):
            col_inds = columnlist
        elif isinstance(columnlist, int):
            col_inds.append(columnlist)
        for col_ind in col_inds:
            rng_col = rng.Columns(col_ind)
            rng_col.FormatConditions.AddIconSetCondition()
            rng_col_fc1 = rng.FormatConditions(1)
            rng_col_fc1.IconSet = excel.ActiveWorkbook.IconSets(c.xl3Triangles)
            rng_col_fc1_ic2 = rng_col_fc1.IconCriteria(2)
            rng_col_fc1_ic3 = rng_col_fc1.IconCriteria(3)
            rng_col_fc1_ic2.Type = c.xlConditionValueNumber
            rng_col_fc1_ic2.Operator = 7
            rng_col_fc1_ic2.Value = 0
            rng_col_fc1_ic3.Type = c.xlConditionValueNumber
            rng_col_fc1_ic3.Operator = 5
            rng_col_fc1_ic3.Value = 0

    def xlConditionValueNumber(self, rng, columnlist):
        col_inds = []
        if isinstance(columnlist, list):
            col_inds = columnlist
        elif isinstance(columnlist, int):
            col_inds.append(columnlist)
        for col_ind in col_inds:
            rng_col = rng.Columns(col_ind)
            rng_col.Style = "Percent"
            rng_col.FormatConditions.AddDatabar()
            rng_col_fc1 = rng.FormatConditions(1)
            rng_col_fc1.MinPoint.Modify(newtype=c.xlConditionValueNumber, newvalue=0)
            rng_col_fc1.MaxPoint.Modify(newtype=c.xlConditionValueNumber, newvalue=1)
            rng_col_fc1.BarColor.Color = 13012579
            rng_col_fc1.BarColor.TintAndShade = 0

    # def hour_style(self):
    #     #无网格线
    #     self.none_gridlines()
    #     # 赋值
    #     rg = self.range
    #     title_rg = rg.Rows(1)
    #     datetime_rg = rg.Rows(2)
    #     cols_rg = rg.Rows(3)
    #     rc=rg.Rows.Count
    #     cc=rg.Columns.Count
    #     data_rg = self.sheet.Range(rg.Cells(4,1),rg.Cells(rc,cc))
    #     # 设置格式
    #     self.title_style(title_rg)
    #     self.datetime_style(datetime_rg)
    #     self.cols_style(cols_rg,False)
    #     self.data_style(data_rg,False)
    #     self.autofit(cols_rg,9)
    #     self.bold(data_rg,'汇总',[1,2,3,4])
    #     self.merge(data_rg,[1,2,3])
    #     self.xl3Triangles(data_rg, 8)
    #     self.xlConditionValueNumber(data_rg, 9)