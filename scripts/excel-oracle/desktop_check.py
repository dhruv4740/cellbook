"""Holds a workbook open in Excel while Cellbook tries Ctrl+S on it, then runs the AutoSave and crash checks.
   Run: npm run build && python scripts/excel-oracle/desktop_check.py"""
import os, shutil, subprocess, win32com.client, pythoncom
src = os.path.abspath('scripts/excel-oracle/out/stress.xlsx')
f = os.path.abspath('scripts/excel-oracle/out/rt/desktop-check.xlsx'); shutil.copy(src, f)
xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False
node = None
try:
    wb = xl.Workbooks.Open(f)
    node = subprocess.Popen(['node', 'scripts/excel-oracle/desktop_check.mjs', f], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True, encoding='utf-8')
    for line in node.stdout:
        if line.strip() == 'LOCKED-DONE': break
        print(line, end='')
    wb.Close(False)
finally:
    xl.Quit(); del xl; pythoncom.CoUninitialize()
if node:
    node.stdin.write('go\n'); node.stdin.flush()
    for line in node.stdout: print(line, end='')
