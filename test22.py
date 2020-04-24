import win32com.client
xlApp = win32com.client.Dispatch('Excel.Application')

xlBook = xlApp.Workbooks.Add()

xlSheet = xlBook.Sheets(1)
xlSheet.Name = "Algoritmos de Busqueda"
xlSheet.Cells(1,1).Value="Secuencial"
xlSheet.Cells(2,1).Value="Binaria"
xlSheet.Cells(1,2).Value="32"
xlSheet.Cells(2,2).Value="32"

chart = xlApp.Charts.Add()
chart.Name= "Grafico "+xlSheet.Name
series = chart.SeriesCollection(1)
series.XValues= xlSheet.Range("A1:A2")
series.Values= xlSheet.Range("B1:B2")
series.Name= "Algoritmos"
chart.Axes()[0].HasMajorGridlines = True