"""Compares Cellbook's function results with real Excel's over func_cases.json. python scripts/excel-oracle/func_oracle.py"""
import json, os, subprocess, win32com.client
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.abspath(os.path.join(HERE, '..', '..'))
cases = json.load(open(os.path.join(HERE, 'func_cases.json'), encoding='utf-8'))
ours = json.loads(subprocess.run('npx vite-node scripts/excel-oracle/func_cb.ts', cwd=ROOT, shell=True, capture_output=True, text=True, check=True).stdout)
ERR = {-2146826288: '#NULL!', -2146826281: '#DIV/0!', -2146826273: '#VALUE!', -2146826265: '#REF!', -2146826259: '#NAME?', -2146826252: '#NUM!', -2146826246: '#N/A'}
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
bad = 0
try:
    wb = xl.Workbooks.Add(); ws = wb.Worksheets(1)
    for r, row in enumerate(cases['data']):
        for c, v in enumerate(row):
            if v is None: continue
            cell = ws.Cells(r + 1, c + 1)
            if isinstance(v, str): cell.NumberFormat = '@'; cell.Value = v
            else: cell.Value = v
    for i, f in enumerate(cases['formulas']):
        ws.Cells(i + 1, 11).Formula2 = '=' + f
    xl.Calculate()
    for i, f in enumerate(cases['formulas']):
        v = ws.Cells(i + 1, 11).Value2
        if isinstance(v, int) and v in ERR: v = ERR[v]
        if isinstance(v, float) and v == int(v): v = int(v)
        o = ours[i]
        if isinstance(o, float) and o == int(o): o = int(o)
        ok = (v == o) or (isinstance(v, float) and isinstance(o, (int, float)) and abs(v - o) < 1e-9) or (v is None and o in (0, '', None))
        if not ok: bad += 1; print('DIFF', f, '| excel:', repr(v), '| cellbook:', repr(o))
    wb.Close(False)
finally:
    xl.Quit()
print('checked', len(cases['formulas']), 'differences', bad)
