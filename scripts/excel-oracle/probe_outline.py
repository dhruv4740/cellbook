import win32com.client
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
def st(ws, n=12):
    return [(r, ws.Rows(r).OutlineLevel, int(ws.Rows(r).Hidden)) for r in range(1, n+1) if ws.Rows(r).OutlineLevel > 1 or ws.Rows(r).Hidden]
try:
    wb = xl.Workbooks.Add(); ws = wb.Worksheets(1)
    for r in range(1, 13): ws.Cells(r, 1).Value = r
    ws.Rows('2:4').Group(); ws.Rows('3:3').Group()
    print('grouped 2:4 + 3', [(r, ws.Rows(r).OutlineLevel) for r in range(1, 8)])
    ws.Outline.ShowLevels(1); print('ShowLevels(1)', st(ws))
    ws.Outline.ShowLevels(2); print('ShowLevels(2)', st(ws))
    ws.Outline.ShowLevels(3); print('ShowLevels(3)', st(ws))
    ws.Rows(5).ShowDetail = False; print('row5 ShowDetail False', st(ws), ws.Rows(5).ShowDetail)
    ws.Rows('2:4').Ungroup(); print('ungroup 2:4 while collapsed', [(r, ws.Rows(r).OutlineLevel, int(ws.Rows(r).Hidden)) for r in range(1, 7)])
    ws.Rows('7:8').Group(); ws.Rows('9:10').Group(); print('adjacent groups', [(r, ws.Rows(r).OutlineLevel) for r in range(6, 12)])
    ws.Rows(11).ShowDetail = False; print('collapse via row 11', [(r, ws.Rows(r).OutlineLevel, int(ws.Rows(r).Hidden)) for r in range(6, 12)])
    ws.Rows('6:6').Group(); print('group 6 next to collapsed', [(r, ws.Rows(r).OutlineLevel, int(ws.Rows(r).Hidden)) for r in range(5, 12)])
    ws.Outline.SummaryRow = 0; print('summary above; ShowDetail row6?', ws.Rows(6).ShowDetail)
    ws.Cells.ClearOutline(); print('clear outline', st(ws))
    # cells range group requires rows/cols spec
    ws.Range('A2:C4').Rows.Group() ; print('range rows group', [(r, ws.Rows(r).OutlineLevel) for r in range(1, 6)])
    ws.Range('A1:A1').Columns.Group(); print('col group', ws.Columns(1).OutlineLevel)
    try:
        ws.Range('A20').Group(); print('single cell Group ok')
    except Exception as e: print('single cell Group err', str(e)[:80])
    # max level
    for i in range(9):
        try: ws.Rows('15:16').Group()
        except Exception as e: print('level limit at', i, str(e)[:60]); break
    print('lvl', ws.Rows(15).OutlineLevel)
finally:
    xl.Quit()
