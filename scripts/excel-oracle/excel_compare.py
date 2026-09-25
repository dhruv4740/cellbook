"""Opens out/formulas.xlsx in real Excel (COM), reads A10:A170 and compares with out/cellbook-values.json.
   Run on Windows with Excel: python scripts/excel-oracle/excel_compare.py   (needs: pip install pywin32)"""
import os, json, datetime, win32com.client, pythoncom
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out')
ERR = {-2146826281: '#DIV/0!', -2146826246: '#N/A', -2146826259: '#NAME?', -2146826288: '#NULL!',
       -2146826252: '#NUM!', -2146826265: '#REF!', -2146826273: '#VALUE!', -2146826245: '#GETTING_DATA',
       -2146826264: '#SPILL!', -2146826263: '#CALC!'}
cb = {r[0]: r for r in json.load(open(os.path.join(OUT, 'cellbook-values.json'), encoding='utf-8'))}
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    wb = xl.Workbooks.Open(os.path.join(OUT, 'formulas.xlsx'))
    ws = wb.Worksheets(1); xl.CalculateFull()
    ex = {}
    for r in range(10, 171):
        c = ws.Cells(r, 1)
        if not c.Formula: continue
        v = c.Value2
        if isinstance(v, int) and v in ERR: v = ERR[v]
        ex[r] = (c.Formula2 if hasattr(c, 'Formula2') else c.Formula, v, c.Text)
    wb.Close(False)
finally:
    xl.Quit(); del xl; pythoncom.CoUninitialize()

def same(e, c):
    if isinstance(e, bool) or isinstance(c, bool): return e == c
    if isinstance(e, (int, float)) and isinstance(c, (int, float)):
        return abs(e - c) <= 1e-9 * max(1, abs(e))
    if e is None: return c in (None, '', 0)
    return str(e) == str(c)
mism = []
for r, (f, v, t) in ex.items():
    c = cb.get(r)
    cv = c[2] if c else '<missing>'
    if not same(v, cv): mism.append((r, f, v, t, cv, c[3] if c else None))
json.dump({'excel': ex, 'mismatches': mism}, open(os.path.join(OUT, 'excel-compare.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1, default=str)
print(len(ex), 'formulas,', len(mism), 'mismatches')
for m in mism: print(m)
