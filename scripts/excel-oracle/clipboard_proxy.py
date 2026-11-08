"""Clipboard check without the clipboard (for machines where the test process can't open it).
   Excel -> Cellbook: Excel saves the block as Unicode text (the same tab-separated text it puts on the
   clipboard) and Cellbook's paste reads it. Cellbook -> Excel: Cellbook's copy HTML is opened by Excel's
   HTML import (the same importer Excel uses for pasted HTML).
   Run: python scripts/excel-oracle/clipboard_proxy.py"""
import os, json, subprocess, win32com.client, pythoncom
OUT = os.path.abspath('scripts/excel-oracle/out'); os.makedirs(OUT, exist_ok=True)
tsv_path = os.path.join(OUT, 'excel-block.txt')
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Add(); ws = wb.Worksheets(1)
    ws.Range('A1').Value = 1234567.891; ws.Range('A1').NumberFormat = '#,##0.00'
    ws.Range('B1').Value = 46268; ws.Range('B1').NumberFormat = 'dd-mm-yyyy'
    ws.Range('C1').Value = 'a, b, c'
    ws.Range('D1').Value = 'Line1' + chr(10) + 'Line2'
    ws.Range('E1').NumberFormat = '@'; ws.Range('E1').Value = '00012'
    ws.Range('F1').Value = -0.125; ws.Range('F1').NumberFormat = '0.0%'
    ws.Range('G1').Value = True
    ws.Range('H1').Formula = '=A1*2'
    ws.Range('I1').Value = 'x"y'
    ws.Range('J1').Value = 5000; ws.Range('J1').NumberFormat = '"₹"#,##0'
    if os.path.exists(tsv_path): os.remove(tsv_path)
    wb.SaveAs(tsv_path, 42); wb.Close(False)

    res = subprocess.run(['npx', 'vite-node', 'scripts/excel-oracle/clipboard_proxy.ts', tsv_path, os.path.join(OUT, 'cellbook-copy.html')], capture_output=True, text=True, encoding='utf-8', shell=True)
    print(res.stdout); print(res.stderr[-2000:])

    wb = xl.Workbooks.Open(os.path.join(OUT, 'cellbook-copy.html')); ws = wb.Worksheets(1)
    print('--- Cellbook copy HTML opened by Excel')
    for c in ws.UsedRange:
        print('  ', c.GetAddress(0, 0), repr(c.Value), c.NumberFormat, repr(c.Text))
    wb.Close(False)
finally:
    xl.Quit(); del xl; pythoncom.CoUninitialize()
