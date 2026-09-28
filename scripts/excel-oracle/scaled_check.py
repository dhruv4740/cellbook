"""Opens samples/private/scaled-N.xlsx (written by Cellbook's fast writer) in real Excel and compares VLOOKUP results and data with Cellbook's."""
import json, os, subprocess, sys, time, win32com.client
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.abspath(os.path.join(HERE, '..', '..'))
n = sys.argv[1] if len(sys.argv) > 1 else '50000'
path = os.path.join(ROOT, 'samples', 'private', 'scaled-%s.xlsx' % n)
ours = json.loads(subprocess.run('npx vite-node scripts/excel-oracle/scaled_cb.ts "%s"' % path, cwd=ROOT, shell=True, capture_output=True, text=True, check=True).stdout)
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    t = time.time(); wb = xl.Workbooks.Open(path); open_s = time.time() - t
    ws = wb.Worksheets(1)
    t = time.time(); xl.Calculate(); calc_s = time.time() - t
    print('excel open %.1fs, calculate %.1fs, used %s, cellbook load %.1fs' % (open_s, calc_s, ws.UsedRange.Address, ours['loadMs'] / 1000))
    bad = 0
    for r, v in ours['rows'].items():
        row = int(r) + 1
        for col, key in ((23, 'w'), (4, 'd'), (5, 'e')):
            x = ws.Cells(row, col).Value2
            o = v[key]
            if isinstance(o, dict): o = o.get('error')
            if isinstance(x, int) and x < -2000000000: x = '#N/A' if x == -2146826246 else x
            ok = x == o or (isinstance(x, (int, float)) and isinstance(o, (int, float)) and abs(x - o) < 1e-9) or (x is None and o in (None, '', 0))
            if not ok:
                bad += 1
                if bad < 10: print('DIFF row', row, key, 'excel', repr(x), 'cellbook', repr(o))
    print('compared', len(ours['rows']) * 3, 'cells, differences:', bad)
    wb.Close(False)
finally:
    xl.Quit()
