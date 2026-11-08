import os, zipfile, hashlib, base64, win32com.client
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'out'); path = os.path.join(OUT, 'probe2.xlsx')
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = False
try:
    wb = xl.Workbooks.Add(); ws = wb.Worksheets(1)
    ws.Protect('abc', True, True, True, False, True, False, False, True, False, True, False, True, True, True)  # positional
    p = ws.Protection
    print({k: getattr(p, k) for k in ['AllowFormattingCells','AllowFormattingColumns','AllowFormattingRows','AllowInsertingColumns','AllowInsertingRows','AllowInsertingHyperlinks','AllowDeletingColumns','AllowDeletingRows','AllowSorting','AllowFiltering','AllowUsingPivotTables']})
    if os.path.exists(path): os.remove(path)
    wb.SaveAs(path, 51); wb.Close(False)
finally: xl.Quit()
x = zipfile.ZipFile(path).read('xl/worksheets/sheet1.xml').decode()
import re; m = re.search(r'<sheetProtection[^>]*>', x).group(0); print(m)
salt = base64.b64decode(re.search(r'saltValue="([^"]+)"', m).group(1)); want = re.search(r'hashValue="([^"]+)"', m).group(1)
h = hashlib.sha512(salt + 'abc'.encode('utf-16-le')).digest()
for i in range(100000): h = hashlib.sha512(h + i.to_bytes(4, 'little')).digest()
print('hash ok', base64.b64encode(h).decode() == want)
