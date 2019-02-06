import win32com.client
import os
import pywintypes
import time
import threading

def _thread(sheet, start, end):
    s = 0
    print(start,end)
    for ii in range(int(start), int(end)):
        r = sheet.Cells(ii,4).Value
        s += r
    return s

if __name__ == '__main__':
    isOpen = False
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        isOpen = True
    except (pywintypes.com_error) as e:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True    
        isOpen = False
        
    excel_file1 = excel.Workbooks.Open(os.path.abspath('weather.xlsx'))
    w_sheet1 = excel_file1.ActiveSheet

    st = time.time()
    s = 0
    r = w_sheet1.Range('C2:D3999').Value
    rr = list(zip(*r))
    print(r[2][1])
    if isOpen is False:
        excel.Quit()