import win32com.client


if __name__ == '__main__':
    excel = win32com.client.Dispatch("Excel.Application")
    excel_file1 = excel.Workbooks.Open(u'weather.xlsx')
    w_sheet1 = excel_file1.ActiveSheet
    first_num = 1
    second_num = 1