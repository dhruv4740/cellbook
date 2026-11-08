import os, win32com.client
HERE = os.path.dirname(os.path.abspath(__file__))
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    wb = xl.Workbooks.Open(os.path.join(HERE, 'out', 'shapes.xlsx'))
    ws = wb.Worksheets(1)
    print('shapes', ws.Shapes.Count)
    for s in ws.Shapes:
        txt = ''
        try: txt = s.TextFrame2.TextRange.Text
        except Exception: pass
        print(' ', s.Name, 'type', s.Type, 'auto', s.AutoShapeType, 'tl', s.TopLeftCell.Address, 'w', round(s.Width / 0.75), 'h', round(s.Height / 0.75), 'text', repr(txt))
finally:
    xl.Quit()
