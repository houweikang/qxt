import xlwings as xw

class MyExcel():
    """
    定义Excel类
    """

    def __init__(self, wb=None):
        if wb:
            self.wb = wb
        else:
            try：
                self.wb = xw.books.active
            except:
                print('不存在打开的工作簿')

        self.app = xw.App(visible=True, add_book=False)

    def sht(self, sht_name, clean = True, DisplayGridlines = False):
        try:
            return self.wb.sheets.add(name=sht_name, after=self.wb.sheets[-1].name)
        except ValueError:
            return self.wb.sheets[sht_name]

        if DisplayGridlines == False:
            self.wb.activate
            self.wb.sheets[sht_name].activate
            self.app.api.ActiveWindow.DisplayGridlines = False

        if clean:
            self.wb.sheets[sht_name].cells.clear()


def main():
    sht = MyExcel().sht('a')

if __name__ == '__main__':
    main()
