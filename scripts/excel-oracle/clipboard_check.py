"""Clipboard between real Excel and the Cellbook desktop app (uses the Windows clipboard).
   Needs: npm run build, npm i --no-save playwright-core, pip install pywin32.
   Run: python scripts/excel-oracle/clipboard_check.py"""
import json, subprocess, time, win32com.client, pythoncom

node = subprocess.Popen(['node', 'scripts/excel-oracle/cellbook_driver.mjs'], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True, encoding='utf-8')
def cb(cmd):
    node.stdin.write(cmd + '\n'); node.stdin.flush()
    return json.loads(node.stdout.readline())
assert cb('select A1') if False else json.loads(node.stdout.readline()).get('ready')

xl = win32com.client.DispatchEx('Excel.Application'); xl.DisplayAlerts = False; xl.Visible = True
try:
    wb = xl.Workbooks.Add(); ws = wb.Worksheets(1)
    ws.Range('A1').Value = 1234.5; ws.Range('A1').NumberFormat = '#,##0.00'
    ws.Range('B1').Value = 46268; ws.Range('B1').NumberFormat = 'dd-mm-yyyy'
    ws.Range('C1').Value = 'a, b, c'
    ws.Range('D1').Value = 'Line1' + chr(10) + 'Line2'
    ws.Range('E1').NumberFormat = '@'; ws.Range('E1').Value = '00012'
    ws.Range('F1').Value = -0.125; ws.Range('F1').NumberFormat = '0.0%'
    ws.Range('G1').Value = True
    ws.Range('H1').Formula = '=A1*2'
    ws.Range('I1').Value = 'x"y'
    ws.Range('A1:I1').Copy(); time.sleep(0.5)

    print('--- 1. Excel -> Cellbook')
    cb('select A1'); cb('key Control+v')
    for c in cb('dump A1:I2')[0] + cb('dump A1:I2')[1][:1]: print('  ', c)

    print('--- 2. Cellbook -> Excel')
    cb('set A5 1234.5'); cb('set B5 3/9/2026'); cb('set C5 a, b, c'); cb('set D5 Line1\\nLine2'); cb("set E5 '00012")
    cb('set F5 12.5%'); cb('set G5 TRUE'); cb('set H5 =A5*2'); cb('set I5 x"y')
    cb('select A5:I5'); cb('key Control+c'); time.sleep(0.5)
    ws.Range('A10').Select(); ws.Paste(); time.sleep(0.5)
    for col in 'ABCDEFGHIJ':
        for r in (10, 11):
            c = ws.Range(f'{col}{r}')
            if c.Value is not None: print('  ', c.GetAddress(0, 0), repr(c.Value), c.NumberFormat, repr(c.Text))

    print('--- 3. Cellbook copy, then Excel copy, then paste in Cellbook')
    cb('select A5'); cb('key Control+c'); time.sleep(0.3)
    ws.Range('C1').Copy(); time.sleep(0.5)
    cb('select K1'); cb('key Control+v')
    print('  ', cb('dump K1:K1')[0][0])
    xl.CutCopyMode = False
    wb.Close(False)
finally:
    xl.Quit(); del xl; pythoncom.CoUninitialize()
    try: cb('quit')
    except Exception: pass
