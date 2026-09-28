"""Screenshots of real Excel's ribbon, one PNG per tab, for comparing Cellbook's ribbon layout.
   python scripts/excel-oracle/ribbon_shots.py [width_px]
   Opens a private, visible Excel window (never touches other Excel windows), sizes it, clicks each ribbon tab with the real mouse and
   grabs the window with PrintWindow. Output: scripts/excel-oracle/out/ribbon/excel-<tab>.png  (the mouse moves while it runs)"""
import ctypes, os, sys, time
ctypes.windll.shcore.SetProcessDpiAwareness(2)  # physical pixels everywhere
import win32api, win32con, win32gui, win32ui, win32com.client
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'out', 'ribbon'); os.makedirs(OUT, exist_ok=True)
WIDTH = int(sys.argv[1]) if len(sys.argv) > 1 else 1752
HEIGHT = 1020
TAB_Y = 81  # tab strip, in window pixels

def grab(hwnd, path):
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r - l, b - t
    dc = win32gui.GetWindowDC(hwnd); mfc = win32ui.CreateDCFromHandle(dc); save = mfc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap(); bmp.CreateCompatibleBitmap(mfc, w, h); save.SelectObject(bmp)
    ctypes.windll.user32.PrintWindow(hwnd, save.GetSafeHdc(), 2)
    info = bmp.GetInfo(); bits = bmp.GetBitmapBits(True)
    Image.frombuffer('RGB', (info['bmWidth'], info['bmHeight']), bits, 'raw', 'BGRX', 0, 1).save(path)
    win32gui.DeleteObject(bmp.GetHandle()); save.DeleteDC(); mfc.DeleteDC(); win32gui.ReleaseDC(hwnd, dc)

def ribbon_window(hwnd):
    found = []
    def cb(h, _):
        if win32gui.GetClassName(h) == 'NetUIHWND': found.append(h)
    win32gui.EnumChildWindows(hwnd, cb, None)
    return found

def click(hwnd, x, y):
    """Posts a mouse click at window coordinates (x, y) to the ribbon control (works without an interactive desktop)."""
    for h in ribbon_window(hwnd):
        l, t, r, b = win32gui.GetWindowRect(h)
        wl, wt, _, _ = win32gui.GetWindowRect(hwnd)
        cx, cy = x + wl - l, y + wt - t
        if 0 <= cx < r - l and 0 <= cy < b - t:
            lp = (cy << 16) | (cx & 0xFFFF)
            win32gui.PostMessage(h, win32con.WM_MOUSEMOVE, 0, lp); time.sleep(0.1)
            win32gui.PostMessage(h, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp); time.sleep(0.08)
            win32gui.PostMessage(h, win32con.WM_LBUTTONUP, 0, lp)
            return True
    return False

xl = win32com.client.DispatchEx('Excel.Application')
try:
    xl.Visible = True; xl.DisplayAlerts = False
    wb = xl.Workbooks.Add()
    hwnd = xl.Hwnd
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.MoveWindow(hwnd, 0, 0, WIDTH, HEIGHT, True)
    try: win32gui.SetForegroundWindow(hwnd)
    except Exception: pass
    time.sleep(1.5)
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    print('window', l, t, r - l, b - t)
    tabs = [('home', 117), ('insert', 185), ('draw', 252), ('pagelayout', 342), ('formulas', 449), ('data', 525), ('review', 596), ('view', 666), ('help', 726)]
    for name, x in tabs:
        print(name, 'clicked' if click(hwnd, x, TAB_Y) else 'no ribbon window')
        time.sleep(1.0)
        grab(hwnd, os.path.join(OUT, 'excel-%s.png' % name))
        print('saved', name)
    wb.Close(False)
finally:
    xl.Quit()
