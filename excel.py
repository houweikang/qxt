import win32com.client
from win32com.client import constants as c  # 旨在直接使用VBA常数
import os

class Excel:
    def __init__(self):
        self.excel = win32com.client.gencache.EnsureDispatch("excel.Application")
        self.excel.Visible = 1
        self.excel.DisplayAlerts = 0
        self.excel.ScreenUpdating = 1
        self.wb=None

    def workbooks(self,name_or_index):
        self.wb=self.excel.Workbooks(name_or_index)
        try:
            return self.wb
        except:
            self.wb=None
            print('Not Existsed Workbooks(%s)' % name_or_index)

    @property
    def activeworkbook(self):
        self.wb=self.excel.ActiveWorkbook
        try:
            return self.wb
        except:
            self.wb=None
            print('Not Existsed ActiveWorkbook!')
        # return xw.books.active

    def workbooks_add(self,name=None,path=None):
        self.wb=self.excel.Workbooks.Add()
        if name:
            MyType='.xlsx'
            if not name.endswith(MyType):
                name=name + MyType
            if not path:
                path=os.path.join(os.path.expanduser("~"), 'Desktop')
            if not os.path.exists(path):
                print('Not exists path(%s),Have changed path to Desktop' % path)
                path=os.path.join(os.path.expanduser("~"), 'Desktop')
            full_name=os.path.join(path, name)
            self.wb.SaveAs(full_name)
            self.wb.Save
        return self.wb

    # def sheets(self,name_or_index):
    #     return self.wb.Sheets(name_or_index)

    def init_wb(self):
        return self.wb

    class sheets:
        def __init__(self,name_or_index):
            wb=Excel().init_wb()
            self.sheet= wb.Sheets(name_or_index)
                # try:
                #     return self.sheet
                # except:
                #     self.sheet=None
                #     print('Not Existsed Sheets(%s)' % name_or_index)
    
    # @property
    # def ActiveSheet(self):
    #     if self.ActiveWorkbook:
    #         self.sheet=self.ActiveWorkbook.ActiveSheet
    #         try:
    #             return self.sheet
    #         except:
    #             print('Not Existsed ActiveSheet')
    #     else:
    #         print('Not exists ActiveWorkbook!')

    # def Sheets_Add(self,name=None,index=None):
    #     if self.wb:
    #         if (not index) or (index > self.wb.Sheets.Count):
    #             index=self.wb.Sheets.Count
    #             self.sheet=self.wb.Sheets.Add(Before = None , After = self.wb.Sheets(index))
    #         elif index > 0:
    #             self.sheet=self.wb.Sheets.Add(Before = self.wb.Sheets(index) , After = None)
    #         if name:
    #             self.sheet.Name = name
    #         try:
    #             return self.sheet
    #         except:
    #             print('Sheets_Add Error!')
    #     else:
    #         print('Not exists Workbook!')
    



def main():
    App=Excel()
    # wb=App.ActiveWorkbook
    wb=App.workbooks(1)
    # wb=App.Workbooks_Add('123','a.daf')
    print(wb.Name)

    # sht=App.sheets(1)
    sht=wb.sheets(1)
    # sht=App.ActiveSheet
    # sht=App.Sheets_Add('abcde')
    print(sht.Name)

if __name__ == "__main__":
    main()