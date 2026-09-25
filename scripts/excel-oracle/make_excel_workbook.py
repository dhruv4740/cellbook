"""Builds out/excel-built.xlsx inside real Excel (COM) with the features employees' files use.
   Run on Windows with Excel: python scripts/excel-oracle/make_excel_workbook.py   (needs pywin32)"""
import os, win32com.client, pythoncom
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out'); os.makedirs(OUT, exist_ok=True)
path = os.path.join(OUT, 'excel-built.xlsx')
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    wb = xl.Workbooks.Add(); ws = wb.Worksheets(1); ws.Name = 'Dispatch'
    rows = [['Code', 'Party', 'Qty', 'Amount', 'Date', 'Status'],
            ['00012', 'माल 5 ट्रक', 10, 1234567.891, 46268, 'Open'],
            ['00034', 'Sharma Traders', 25, 98765.4, 46023, 'Closed'],
            ['00056', 'Gupta & Sons', 0, 250000, 46100, 'Open'],
            ['00078', 'Sharma Traders', 40, 5000.5, 46200, 'Pending'],
            ['00090', 'राम लॉजिस्टिक्स', 5, 12.25, 45700, 'Open']]
    ws.Range('A:A').NumberFormat = '@'
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row, 1): ws.Cells(r, c).Value = v
    ws.Range('D2:D6').NumberFormat = '#,##,##0.00'
    ws.Range('E2:E6').NumberFormat = 'dd-mm-yyyy'
    lo = ws.ListObjects.Add(1, ws.Range('A1:F6'), None, 1); lo.Name = 'Orders'
    wb.Names.Add('Rate', '=Dispatch!$I$1'); ws.Range('I1').Value = 0.18
    ws.Names.Add('LocalQty', '=Dispatch!$C$2:$C$6')
    f = {'H2': '=XLOOKUP("00056",A2:A6,B2:B6)', 'H3': '=IFS(C2>20,"big",C2>5,"mid",TRUE,"small")',
         'H4': '=LET(x,SUM(C2:C6),x*Rate)', 'H5': '=SUM(LocalQty)', 'H6': '=SUM(Orders[Qty])',
         'H7': '=ROWS(UNIQUE(B2:B6))', 'H8': '=FILTER(A2:A6,F2:F6="Open")'}
    for a, v in f.items(): ws.Range(a).Formula2 = v
    ws.Range('K1').Value = 'note here'; ws.Range('K1').AddComment('Plain note')
    ws.Range('K2').Value = 'thread here'; ws.Range('K2').AddCommentThreaded('Threaded comment')
    ws.Range('B2').WrapText = False
    ws.Range('F2:F6').Validation.Delete(); ws.Range('F2:F6').Validation.Add(3, 1, 1, 'Open,Closed,Pending')
    ws.Range('C2:C6').FormatConditions.Add(1, 5, '=20').Interior.Color = 0xCCC7FF
    ws.Range('D2:D6').FormatConditions.AddColorScale(3)
    ws.Range('C2:C6').FormatConditions.AddDatabar()
    ws.Range('M1:N2').Merge(); ws.Range('M1').Value = 'Merged'
    ws.Range('P1').Value = 'Line1\nLine2'.replace('\n', chr(10)); ws.Range('P1').WrapText = True
    ch = ws.ChartObjects().Add(400, 150, 300, 200).Chart; ch.SetSourceData(ws.Range('B1:C6')); ch.ChartType = 51
    ws.Columns('J').Hidden = True; ws.Rows(9).Hidden = True
    lo.Range.AutoFilter(6, 'Open')
    ws.Activate(); xl.ActiveWindow.SplitRow = 1; xl.ActiveWindow.SplitColumn = 1; xl.ActiveWindow.FreezePanes = True
    if os.path.exists(path): os.remove(path)
    wb.SaveAs(path, 51); wb.Close(False)
    print('wrote', path)
finally:
    xl.Quit(); del xl; pythoncom.CoUninitialize()
