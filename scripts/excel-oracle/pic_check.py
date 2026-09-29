import os, win32com.client
HERE = os.path.dirname(os.path.abspath(__file__))
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    wb = xl.Workbooks.Open(os.path.join(HERE, 'out', 'pic.xlsx'))
    ws = wb.Worksheets(1)
    print('shapes', ws.Shapes.Count)
    for s in ws.Shapes:
        print(' ', s.Name, 'type', s.Type, 'tl', s.TopLeftCell.Address, 'w', round(s.Width / 0.75), 'h', round(s.Height / 0.75), 'alt', s.AlternativeText)
finally:
    xl.Quit()
