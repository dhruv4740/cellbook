import sys, os, json, win32com.client
sys.path.insert(0, os.path.dirname(__file__))
from feature_rt import with_excel, OUT
def fn(xl, wb):
    for k in ['default', 'letterLandscape', 'fit1xAuto']:
        ws = wb.Worksheets(k); ps = ws.PageSetup
        ws.Activate(); xl.ActiveWindow.View = 2  # page break preview computes HPageBreaks reliably
        h = [ws.HPageBreaks(i).Location.Row for i in range(1, ws.HPageBreaks.Count + 1)]
        v = [ws.VPageBreaks(i).Location.Column for i in range(1, ws.VPageBreaks.Count + 1)]
        print(k, 'paper', ps.PaperSize, 'orient', ps.Orientation, 'zoom', ps.Zoom, 'pages', ps.Pages.Count, 'H', h, 'V', v,
              'defRowH', ws.StandardHeight, 'colW', ws.Columns(1).Width, 'colWchars', ws.Columns(1).ColumnWidth, 'L/R', ps.LeftMargin, ps.RightMargin, 'T/B', ps.TopMargin, ps.BottomMargin)
with_excel(os.path.join(OUT, 'rt-pages.xlsx'), fn)
