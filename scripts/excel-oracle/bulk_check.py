"""Opens out/bulk.xlsx (written by bulk_write.ts) in real Excel and compares cell contents / formats to the source."""
import json, os, win32com.client
HERE = os.path.dirname(os.path.abspath(__file__))
src = json.load(open(os.path.join(HERE, 'out', 'bulk.json'), encoding='utf-8'))
cells = src['cells']
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
bad = 0
def rgb(hex_): h = hex_.lstrip('#'); return int(h[4:6], 16) * 65536 + int(h[2:4], 16) * 256 + int(h[0:2], 16)
try:
    wb = xl.Workbooks.Open(os.path.join(HERE, 'out', 'bulk.xlsx'))
    ws = wb.Worksheets('Big')
    print('sheets', [s.Name for s in wb.Worksheets], 'used', ws.UsedRange.Address)
    import random
    keys = list(cells.keys()); random.seed(1); sample = random.sample(keys, 400) + ['0,7', '1,7', '2,7', '3,7', '10,8', '0,0', '0,1']
    for k in sample:
        r, c = map(int, k.split(',')); cell = cells[k]; x = ws.Cells(r + 1, c + 1)
        raw = cell.get('raw'); st = cell.get('style') or {}
        problems = []
        if raw is not None:
            if raw.startswith('='):
                if x.Formula.replace(' ', '') != raw.replace(' ', ''): problems.append(('formula', x.Formula, raw))
            elif raw.startswith("'"):
                if x.Value2 != raw[1:]: problems.append(('text', x.Value2, raw))
            elif raw in ('TRUE', 'FALSE'):
                if x.Value2 != (raw == 'TRUE'): problems.append(('bool', x.Value2, raw))
            else:
                try:
                    if abs(float(x.Value2) - float(raw)) > 1e-9: problems.append(('num', x.Value2, raw))
                except Exception: 
                    if x.Value2 != raw: problems.append(('str', x.Value2, raw))
        if st.get('bold') and not x.Font.Bold: problems.append(('bold',))
        if st.get('italic') and not x.Font.Italic: problems.append(('italic',))
        if st.get('strike') and not x.Font.Strikethrough: problems.append(('strike',))
        if st.get('color') and x.Font.Color != rgb(st['color']): problems.append(('color', x.Font.Color))
        if st.get('fill') and x.Interior.Color != rgb(st['fill']): problems.append(('fill', x.Interior.Color))
        if st.get('fontFamily') and x.Font.Name != st['fontFamily']: problems.append(('font', x.Font.Name))
        if st.get('fontSize') and x.Font.Size != st['fontSize']: problems.append(('size', x.Font.Size))
        BS = chr(92)
        if st.get('numFmt') and x.NumberFormat.replace(BS, '') != st['numFmt'].replace(BS, ''): problems.append(('numFmt', x.NumberFormat, st['numFmt']))
        if st.get('hAlign') == 'center' and x.HorizontalAlignment != -4108: problems.append(('hAlign', x.HorizontalAlignment))
        if st.get('wrap') and not x.WrapText: problems.append(('wrap',))
        if st.get('locked') is False and x.Locked: problems.append(('locked',))
        if st.get('indent') and x.IndentLevel != st['indent']: problems.append(('indent', x.IndentLevel))
        b = st.get('borders')
        if b and b.get('bottom') and x.Borders(9).LineStyle != -4119: problems.append(('border-bottom', x.Borders(9).LineStyle))
        if problems: bad += 1; print('CELL', k, problems)
    print('merged?', ws.Range('I11').MergeCells, 'note:', ws.Range('A5').Comment.Text() if ws.Range('A5').Comment else None, 'hidden row 6:', ws.Rows(6).Hidden, 'row3 height', ws.Rows(3).RowHeight, 'col A width', ws.Columns(1).ColumnWidth)
    print('small A1:', wb.Worksheets('Small').Range('A1').Value2)
    wb.Close(False)
finally:
    xl.Quit()
print('bad cells:', bad)
