import os, win32com.client
HERE = os.path.dirname(os.path.abspath(__file__))
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    wb = xl.Workbooks.Open(os.path.join(HERE, 'out', 'scen.xlsx'))
    ws = wb.Worksheets(1)
    print('scenarios', ws.Scenarios().Count)
    for s in ws.Scenarios():
        try: cm = s.Comment
        except Exception: cm = None
        print(' ', s.Name, 'changing', s.ChangingCells.Address, 'values', list(s.Values), 'comment', cm)
    ws.Scenarios('High').Show()
    print('A3 after Show High =', ws.Range('A3').Value2)
finally:
    xl.Quit()
