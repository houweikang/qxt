import win32com.client
from win32com.client import constants as c  # 旨在直接使用VBA常数
import os

excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
excel.Visible = 1
excel.DisplayAlerts = 0
excel.ScreenUpdating = 1

def WorkBooks(name_or_index):
    try:
        return excel.Workbooks(name_or_index)
    except:
        print('Error!')
    
def ActiveWorkBook():
    return excel.ActiveWorkbook
    # return xw.books.active

def WorkBooks_Add(name,path):
    wb=excel.Workbooks.Add()
    MyType='.xlsx'
    if not name.endswith(MyType):
        name=name + MyType
    full_name=os.path.join(path, name)
    wb.SaveAs(full_name)
    wb.Save()
    return wb

def main():
    # wb=WorkBooks_Add('a','.')
    wb=ActiveWorkBook()
    # wb=WorkBooks(1)
    print(wb.Name)

if __name__ == "__main__":
    main()