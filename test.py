import win32com.client
class easyExcel:
    def __init__(self,  visible=0, filename=None):
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        self.xlApp.Visible = visible
        if filename:
            try:
                self.xlBook = self.xlApp.Workbooks[filename]
            except:
                print('工作簿不存在！')

    def open(self, path=None):
        if path:
            try:
                self.xlBook = self.xlApp.Workbooks.Open(path)
            except:
                print('Excel文件不存在')

    def add_wb(self, wb_name=None):
        if wb_name:
            self.xlBook = self.xlApp.Workbooks.Add(wb_name)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
        return self.xlBook

    def active_wb(self):
        self.xlBook = self.xlApp.Activeworkbook()

    def active_sht(self):
        self.sht = self.active_wb().Activesheet()

    # def add_sht(self):


    def save(self, newfilename=None):
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self, save_yes=False):
        self.xlBook.Close(SaveChanges=save_yes)
        del self.xlApp

    def sht(self,sht_name=None):
        if sht_name:
            return self.xlBook.Worksheets(sht_name)
        else:
            return self.xlBook.Activesheet()


    def getCell(self, sheet, row, col):
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col)

    def setCell(self, sheet, row, col, value):
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def getRange(self, sheet, row1, col1, row2, col2):
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2))

    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):
        "Insert a picture in sheet"
        sht = self.xlBook.Worksheets(sheet)
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

    def cpSheet(self, before):
        "copy sheet"
        shts = self.xlBook.Worksheets
        shts(1).Copy(None,shts(1))

#下面是一些测试代码。
if __name__ == "__main__":
    # PNFILE = r'c:\screenshot.bmp'
    app = easyExcel(visible=1)
    wb=app.active_wb()
    print(wb.name)
    sht=wb.active_sht()
    print(sht.name)
    app.close()
