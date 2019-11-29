import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.visible = True
# wb = excel.Workbooks.Open(r'c:\Users\Administrator\Desktop\2019-11-25.xlsx')
excel.Windows[excel.Activeworkbook.name].DisplayGridlines = False
excel.Activeworkbook.save
excel.quit()