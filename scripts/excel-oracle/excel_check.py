"""Opens Cellbook-saved .xlsx files in real Excel (COM) and checks them; re-saves each as excel-<name>.xlsx.
   Usage: python scripts/excel-oracle/excel_check.py file1.xlsx [file2.xlsx ...]   (needs pywin32, openpyxl)"""
import os, sys, glob, time, openpyxl, win32com.client, pythoncom
ERR = {-2146826281: '#DIV/0!', -2146826246: '#N/A', -2146826259: '#NAME?', -2146826288: '#NULL!',
       -2146826252: '#NUM!', -2146826265: '#REF!', -2146826273: '#VALUE!', -2146826264: '#SPILL!', -2146826263: '#CALC!'}
TMP = os.environ['TEMP']
def same(a, b):
    if isinstance(a, bool) or isinstance(b, bool): return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)): return abs(a - b) <= 1e-9 * max(1, abs(a))
    if a in (None, '') and b in (None, ''): return True
    return str(a) == str(b)
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    for path in map(os.path.abspath, sys.argv[1:]):
        print('==', os.path.basename(path))
        cached = openpyxl.load_workbook(path, data_only=True)
        before = set(glob.glob(os.path.join(TMP, 'error*.xml')))
        wb = xl.Workbooks.Open(path)
        repair = set(glob.glob(os.path.join(TMP, 'error*.xml'))) - before
        caption = xl.ActiveWindow.Caption if xl.Windows.Count else ''
        print('  repaired:', bool(repair) or 'Repaired' in str(caption), [open(f, encoding='utf-8', errors='ignore').read()[:600] for f in repair])
        xl.CalculateFull()
        bad = []; errs = []
        for ws in wb.Worksheets:
            ur = ws.UsedRange
            for c in ur:
                if not c.HasFormula: continue
                v = c.Value2
                if isinstance(v, int) and v in ERR: v = ERR[v]
                if v in ('#NAME?', '#REF!'): errs.append((ws.Name, c.GetAddress(0, 0), c.Formula, v))
                cv = cached[ws.Name][c.GetAddress(0, 0)].value if ws.Name in cached.sheetnames else None
                if not same(v, cv): bad.append((ws.Name, c.GetAddress(0, 0), c.Formula, 'excel=', v, 'cached=', cv))
            print(f'  {ws.Name}: merges/freeze', ws.Parent.Windows(1).FreezePanes if ws.Index == 1 else '',
                  'comments', ws.Comments.Count, 'threaded', ws.CommentsThreaded.Count, 'links', ws.Hyperlinks.Count,
                  'cf', ws.Cells.FormatConditions.Count, 'autofilter', ws.AutoFilterMode, 'listobjects', ws.ListObjects.Count,
                  'charts', ws.ChartObjects().Count, 'shapes', ws.Shapes.Count)
        print('  names:', [(n.Name, n.RefersTo) for n in wb.Names])
        print('  #NAME?/#REF!:', errs)
        print('  excel != cached:', len(bad)); [print('   ', b) for b in bad[:20]]
        out = os.path.join(os.path.dirname(path), 'excel-' + os.path.basename(path))
        if os.path.exists(out): os.remove(out)
        wb.SaveAs(out, 51); wb.Close(False)
finally:
    xl.Quit(); del xl; pythoncom.CoUninitialize()
