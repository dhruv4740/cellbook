"""Cellbook writes a workbook -> a private Excel instance opens it (no repair) and reports what it sees.
   python feature_rt.py <case>   (cases below). Uses feature_io.ts through vite-node; never touches other Excel windows."""
import json, os, subprocess, sys, win32com.client
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.abspath(os.path.join(HERE, '..', '..')); OUT = os.path.join(HERE, 'out'); os.makedirs(OUT, exist_ok=True)

def cellbook_write(data, name):
    j = os.path.join(OUT, name + '.json'); x = os.path.join(OUT, name + '.xlsx')
    json.dump(data, open(j, 'w', encoding='utf-8'))
    subprocess.run('npx vite-node scripts/excel-oracle/feature_io.ts write "%s" "%s"' % (j, x), cwd=ROOT, shell=True, check=True)
    return x

def with_excel(path, fn):
    xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
    try:
        wb = xl.Workbooks.Open(path)
        try: return fn(xl, wb)
        finally: wb.Close(False)
    finally: xl.Quit()

ALLOW_OFF = dict(selectLockedCells=True, selectUnlockedCells=True, formatCells=False, formatColumns=False, formatRows=False, insertColumns=False, insertRows=False, insertHyperlinks=False, deleteColumns=False, deleteRows=False, sort=False, autoFilter=False, pivotTables=False, objects=False, scenarios=False)

def case_protection():
    # password hashes are produced by Cellbook's own code (model/protection.ts) -> generate through a tiny vite-node call
    h = json.loads(subprocess.run('npx vite-node scripts/excel-oracle/hash_pw.ts abc', cwd=ROOT, shell=True, capture_output=True, text=True).stdout)
    hb = json.loads(subprocess.run('npx vite-node scripts/excel-oracle/hash_pw.ts xyz', cwd=ROOT, shell=True, capture_output=True, text=True).stdout)
    data = {'sheets': [{'name': 'P', 'cells': {'0,0': {'raw': '1'}, '0,1': {'raw': '2', 'style': {'locked': False}}, '0,2': {'raw': '=A1+1', 'style': {'hidden': True}}},
                        'layout': {'protection': {'password': h, 'allow': dict(ALLOW_OFF, formatCells=True, sort=True, insertRows=True, selectUnlockedCells=True)}}},
                       {'name': 'Q', 'cells': {}, 'layout': {'protection': {'allow': dict(ALLOW_OFF, selectLockedCells=False)}}}],
            'protection': {'lockStructure': True, 'lockWindows': False, 'password': hb}, 'calcMode': 'manual'}
    def fn(xl, wb):
        ws = wb.Worksheets('P'); p = ws.Protection
        res = dict(protectContents=ws.ProtectContents, fmtCells=p.AllowFormattingCells, sort=p.AllowSorting, insRows=p.AllowInsertingRows, fmtCols=p.AllowFormattingColumns,
                   b2locked=ws.Range('B2').Locked, a1locked=ws.Range('A1').Locked, c1hidden=ws.Range('C1').FormulaHidden, bookStruct=wb.ProtectStructure, calc=xl.Calculation)
        try: ws.Unprotect('wrong'); res['wrongPwAccepted'] = True
        except Exception: res['wrongPwAccepted'] = False
        ws.Unprotect('abc'); res['unprotectOk'] = not ws.ProtectContents
        wb.Unprotect('xyz'); res['bookUnprotectOk'] = not wb.ProtectStructure
        q = wb.Worksheets('Q'); res['qEnableSelection'] = q.EnableSelection  # xlNoSelection=-4142, xlUnlockedCells=1, 0=all
        return res
    print(json.dumps(with_excel(cellbook_write(data, 'rt-protection'), fn), indent=1))

def case_outline():
    cells = {'%d,0' % r: {'raw': str(r + 1)} for r in range(10)}
    lay = {'rowLevels': {'1': 1, '2': 2, '3': 1}, 'hiddenRows': {'1': True, '2': True, '3': True}, 'colLevels': {'2': 1, '3': 1}, 'hiddenCols': {'2': True, '3': True},
           'colWidths': {}, 'rowHeights': {}, 'merges': [], 'frozen': {'rows': 0, 'cols': 0}, 'showGridlines': True, 'showHeadings': True, 'zoom': 100, 'notes': {}, 'comments': {}}
    data = {'sheets': [{'name': 'O', 'cells': cells, 'layout': lay},
                       {'name': 'Above', 'cells': cells, 'layout': dict(lay, outline={'summaryBelow': False}, hiddenRows={}, hiddenCols={})}]}
    def fn(xl, wb):
        ws = wb.Worksheets('O')
        res = dict(levels=[ws.Rows(r).OutlineLevel for r in range(1, 7)], hidden=[int(ws.Rows(r).Hidden) for r in range(1, 7)],
                   colLevels=[ws.Columns(c).OutlineLevel for c in range(1, 6)], colHidden=[int(ws.Columns(c).Hidden) for c in range(1, 6)],
                   summaryRow=ws.Outline.SummaryRow, showDetailRow5=ws.Rows(5).ShowDetail, showDetailCol5=ws.Columns(5).ShowDetail)
        a = wb.Worksheets('Above'); res['aboveSummaryRow'] = a.Outline.SummaryRow
        ws.Outline.ShowLevels(3, 3); res['afterShowAll'] = [int(ws.Rows(r).Hidden) for r in range(1, 7)]
        return res
    print(json.dumps(with_excel(cellbook_write(data, 'rt-outline'), fn)))

if __name__ == '__main__':
    globals()['case_' + sys.argv[1]]()
