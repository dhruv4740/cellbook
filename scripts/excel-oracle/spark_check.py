import os, win32com.client
HERE = os.path.dirname(os.path.abspath(__file__))
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    wb = xl.Workbooks.Open(os.path.join(HERE, 'out', 'spark.xlsx'))
    ws = wb.Worksheets(1)
    print('groups', ws.Cells.SparklineGroups.Count)
    for g in ws.Cells.SparklineGroups:
        print(' type', g.Type, 'count', g.Count, 'color', g.SeriesColor.Color, 'loc', g.Location.Address, 'src', g.SourceData)
finally:
    xl.Quit()
