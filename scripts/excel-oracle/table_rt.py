"""Excel builds a workbook with a table -> Cellbook opens and saves it -> Excel reopens: is the table still a ListObject?"""
import os, subprocess, win32com.client
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.abspath(os.path.join(HERE, '..', '..'))
out = os.path.join(HERE, 'out'); os.makedirs(out, exist_ok=True)
src, dst = os.path.join(out, 'table_src.xlsx'), os.path.join(out, 'table_rt.xlsx')
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    wb = xl.Workbooks.Add(); ws = wb.Worksheets(1)
    ws.Range('A1:C1').Value = ('Item', 'Qty', 'Price')
    for i in range(2, 7):
        ws.Cells(i, 1).Value = f'x{i}'; ws.Cells(i, 2).Value = i; ws.Cells(i, 3).Value = i * 1.5
    lo = ws.ListObjects.Add(1, ws.Range('A1:C6'), None, 1); lo.Name = 'Sales'; lo.TableStyle = 'TableStyleMedium9'
    ws.Range('E2').Formula = '=SUM(Sales[Qty])'
    if os.path.exists(src): os.remove(src)
    wb.SaveAs(src, 51); wb.Close(False)
finally:
    xl.Quit()
js = f"""
import fs from 'fs'
import {{ readXlsx, writeXlsx }} from '../../src/renderer/src/io/xlsx'
const data = await readXlsx(new Uint8Array(fs.readFileSync({src!r})))
console.log('tables read:', JSON.stringify(data.tables ?? null).slice(0, 300))
fs.writeFileSync({dst!r}, await writeXlsx(data))
"""
open(os.path.join(HERE, 'out', 'table_rt.ts'), 'w').write(js.replace('../../src', '../../../src'))
r = subprocess.run('npx vite-node scripts/excel-oracle/out/table_rt.ts', cwd=ROOT, shell=True, capture_output=True, text=True)
print(r.stdout.strip(), r.stderr.strip()[:300])
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    wb = xl.Workbooks.Open(dst); ws = wb.Worksheets(1)
    print('list objects after round trip:', ws.ListObjects.Count, [(l.Name, l.Range.Address) for l in ws.ListObjects])
    lo = ws.ListObjects(1); print('style', lo.TableStyle.Name if lo.TableStyle else None, 'stripes', lo.ShowTableStyleRowStripes, 'filter buttons', lo.ShowAutoFilter, 'cols', [c.Name for c in lo.ListColumns])
    print('E2 formula:', ws.Range('E2').Formula, '=', ws.Range('E2').Value2)
    wb.Close(False)
finally:
    xl.Quit()
