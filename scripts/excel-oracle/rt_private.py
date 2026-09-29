"""Opens the original despatch workbook and Cellbook's round-tripped copies in real Excel and compares them."""
import os, subprocess, win32com.client
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.abspath(os.path.join(HERE, '..', '..'))
subprocess.run('npx vite-node scripts/excel-oracle/rt_private.ts', cwd=ROOT, shell=True, check=True)
orig = os.path.join(ROOT, 'samples', 'private', 'DESPATCH DHRUV.xlsx')
rt = os.path.join(HERE, 'out', 'rt-private.xlsx'); rts = os.path.join(HERE, 'out', 'rt-private-sorted.xlsx')
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
def norm(v):
    if isinstance(v, float) and v == int(v): return int(v)
    return v
try:
    a = xl.Workbooks.Open(orig); b = xl.Workbooks.Open(rt); c = xl.Workbooks.Open(rts)
    wa, wb_, wc = a.Worksheets(1), b.Worksheets(1), c.Worksheets(1)
    ua, ub = wa.UsedRange, wb_.UsedRange
    print('orig used', ua.Address, '| roundtrip used', ub.Address, '| sheet names', a.Worksheets(1).Name, b.Worksheets(1).Name)
    va = ua.Value2; vb = ub.Value2
    diffs = 0
    for r in range(1, ua.Rows.Count + 1):
        for cc in range(1, ua.Columns.Count + 1):
            x, y = norm(va[r - 1][cc - 1]), norm(vb[r - 1][cc - 1])
            if x != y and not (isinstance(x, float) and isinstance(y, float) and abs(x - y) < 1e-9):
                diffs += 1
                if diffs < 8: print('VALUE DIFF', r, cc, repr(x), repr(y))
    print('cell value differences:', diffs, 'of', ua.Rows.Count * ua.Columns.Count)
    fd = 0
    import random; random.seed(3)
    for _ in range(600):
        r, cc = random.randint(1, ua.Rows.Count), random.randint(1, ua.Columns.Count)
        p, q = wa.Cells(r, cc), wb_.Cells(r, cc)
        for name, f in (('numfmt', lambda z: z.NumberFormat), ('bold', lambda z: z.Font.Bold), ('color', lambda z: z.Font.Color), ('fill', lambda z: z.Interior.Color), ('halign', lambda z: z.HorizontalAlignment), ('size', lambda z: z.Font.Size), ('font', lambda z: z.Font.Name)):
            if f(p) != f(q):
                fd += 1
                if fd < 8: print('FORMAT DIFF', name, r, cc, repr(f(p)), repr(f(q)))
    print('format differences in 600 sampled cells:', fd)
    print('col widths equal:', all(abs(wa.Columns(i).ColumnWidth - wb_.Columns(i).ColumnWidth) < 0.6 for i in range(1, ua.Columns.Count + 1)))
    print('freeze/zoom:', a.Windows(1).FreezePanes, b.Windows(1).FreezePanes, a.Windows(1).Zoom, b.Windows(1).Zoom)
    print('sorted+filtered copy: autofilter on =', wc.AutoFilterMode, ', visible rows =', wc.AutoFilter.Range.SpecialCells(12).Rows.Count if wc.AutoFilterMode else None, ', first data source cell =', wc.Cells(2, 5).Value2)
    a.Close(False); b.Close(False); c.Close(False)
finally:
    xl.Quit()
