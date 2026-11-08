"""Has a private Excel instance write protection/outline/header-footer/print-titles/manual-calc, so we can copy exact XML."""
import os, zipfile, win32com.client
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out'); os.makedirs(OUT, exist_ok=True)
path = os.path.join(OUT, 'probe-features.xlsx')
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    wb = xl.Workbooks.Add(); ws = wb.Worksheets(1); ws.Name = 'P'
    for r in range(1, 9): ws.Cells(r, 1).Value = r; ws.Cells(r, 2).Value = r * 2
    ws.Range('B2:B3').Locked = False; ws.Range('B4').FormulaHidden = True
    ws.Rows('2:4').Group(); ws.Rows('3:3').Group(); ws.Columns('D:E').Group(); ws.Outline.ShowLevels(1, 1)
    ws.Outline.SummaryRow = 0  # xlAbove
    ws.PageSetup.LeftHeader = 'L'; ws.PageSetup.CenterHeader = '&D &"Arial,Bold"&12Title'; ws.PageSetup.RightFooter = 'Page &P of &N'
    ws.PageSetup.CenterFooter = '&F - &A'; ws.PageSetup.PrintTitleRows = '$1:$2'; ws.PageSetup.PrintTitleColumns = '$A:$A'
    ws.PageSetup.DifferentFirstPageHeaderFooter = True; ws.PageSetup.FirstPage.CenterHeader.Text = 'First'
    ws.PageSetup.Zoom = False; ws.PageSetup.FitToPagesWide = 1; ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.CenterHorizontally = True; ws.PageSetup.PrintArea = '$A$1:$B$8'
    ws.PageSetup.LeftMargin = 36; ws.PageSetup.TopMargin = 54; ws.PageSetup.HeaderMargin = 20
    ws.PageSetup.Order = 2; ws.PageSetup.BlackAndWhite = True
    ws.Protect(Password='abc', DrawingObjects=True, Contents=True, Scenarios=True, AllowFormattingCells=True, AllowSorting=True, AllowFiltering=True, AllowInsertingRows=True)
    ws2 = wb.Worksheets.Add(After=ws); ws2.Name = 'Q'; ws2.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
    ws2.EnableSelection = 1  # xlNoRestrictions=0, xlUnlockedCells=1
    wb.Protect(Password='xyz', Structure=True, Windows=False)
    xl.Calculation = -4135; xl.CalculateBeforeSave = False
    if os.path.exists(path): os.remove(path)
    wb.SaveAs(path, 51)
    wb.Close(False)
finally:
    xl.Quit()
z = zipfile.ZipFile(path)
for n in ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']:
    print('=====', n); print(z.read(n).decode()[:6000])
