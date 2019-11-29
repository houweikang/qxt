import win32com.client


excel = win32com.client.Dispatch('excel.application')
excel.visible = 1  # 显示Excel
# excel.DisplayAlerts = 0  # 关闭系统警告
# excel.ScreenUpdating = 0  # 关闭屏幕刷新

# 打开Excel文件
workbook = excel.workbooks.Add()
worksheet = workbook.worksheets[0]
worksheet.Cells(1,1).Resize(1,2).Value = [1, 1]
#     [
#     [11, 12, 13, 14],
#     [21, 22, 23, 24],
#     [31, 32, 33, 34],
#     [41, 42, 43, 44]
# ]

# 其他操作代码
# ...

# 关闭工作簿，不保存(若保存，使用True即可)
# workbook.Close(True)
#
# # 退出Excel
# excel.Quit()
