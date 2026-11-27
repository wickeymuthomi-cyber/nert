"""
excel_typer.py - types the data of an Excel file into Excel at human speed,
character by character, cell by cell, exactly as it appears in the source.

The text is entered with real keystrokes, so each cell goes into edit mode and
you see the blinking text caret moving along as every character appears, just
like typing by hand. The green cell cursor and the mouse pointer move to each
cell first.

Moving the mouse does NOT interrupt it (only the keyboard is used for typing).
If you switch to another window, typing pauses by itself and continues as
soon as Excel is in front again, so no characters are lost in other programs.
Do not type on the keyboard yourself while it runs.

Setup (Windows, Excel installed):
    py -m pip install pywin32 openpyxl
Run:
    py excel_typer.py
Put customer_meter_data.xlsx in the same folder as this script.

Output goes to customer_meter_data_typed.xlsx next to the script. It is saved
every few minutes. If anything interrupts the run, start it again: it resumes
from the last saved row. Delete the .progress file to start over.

Excel and Python must run at the same level: if Excel was started "as
administrator", start VS Code as administrator too, or the keystrokes are
blocked by Windows.
"""
import ctypes
import datetime
import os
import random
import sys
import time
from ctypes import wintypes

import openpyxl
import pythoncom  # type: ignore[reportMissingModuleSource]
import win32com.client  # type: ignore[reportMissingModuleSource]

# ---------------------------------------------------------------- settings
HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "customer_meter_data.xlsx")
SHEET = "Data"                 # sheet to type, or None for the first sheet
OUTPUT = os.path.join(HERE, "customer_meter_data_typed.xlsx")

# Typing speed: random seconds between characters
CHAR_DELAY = (0.3, 0.6)
HESITATE_CHANCE = 0.03         # chance of a short thinking pause after a character
HESITATE = (0.4, 1.0)
CELL_PAUSE = (2, 2.5)        # moving on to the next cell
ROW_PAUSE = (3, 3.6)         # finishing a row

PACED = False                  # True = stretch the whole job to TARGET_HOURS
TARGET_HOURS = 7.0             # only used when PACED = True
BREAK_CHANCE = 0.03            # chance of a longer break after a row (PACED)
SAVE_MINUTES = 5
RETRIES = 15                   # attempts per step when Excel is busy
MOVE_MOUSE = True              # the mouse pointer glides to each cell as it is typed
MOUSE_GLIDE = 0.35             # seconds the pointer takes to glide to the next cell

PROGRESS = OUTPUT + ".progress"
APP = None                     # the Excel application, set in main()
_warned = set()


# ----------------------------------------------------------------- helpers
def log(msg):
    print(time.strftime("[%H:%M:%S] ") + msg, flush=True)


def warn_once(key, msg):
    if key not in _warned:
        _warned.add(key)
        log(msg)


def keep_awake():
    """Ask Windows not to sleep while this script runs."""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x1 | 0x2)
    except Exception:
        log("Could not set keep-awake; disable sleep in Windows power settings.")


def com(fn, *args):
    """Run an Excel call, retrying while Excel is busy."""
    for attempt in range(1, RETRIES + 1):
        try:
            return fn(*args)
        except pythoncom.com_error as e:
            if attempt == RETRIES:
                raise
            log(f"Excel busy ({e.hresult & 0xFFFFFFFF:#x}), retry {attempt}/{RETRIES}")
            time.sleep(2 * attempt)


# ------------------------------------------------- keyboard / mouse (Windows)
class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG),
                ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("dwExtraInfo", ctypes.c_size_t)]


class _KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_size_t)]


class _HARDWAREINPUT(ctypes.Structure):
    _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD)]


class _INPUTUNION(ctypes.Union):
    _fields_ = [("mi", _MOUSEINPUT), ("ki", _KEYBDINPUT), ("hi", _HARDWAREINPUT)]


class _INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUTUNION)]


class _PT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


KEYEVENTF_KEYUP, KEYEVENTF_UNICODE = 0x2, 0x4
VK_TAB, VK_RETURN = 0x09, 0x0D


def _send(events):
    arr = (_INPUT * len(events))(*events)
    ctypes.windll.user32.SendInput(len(events), arr, ctypes.sizeof(_INPUT))


def _key(vk=0, scan=0, flags=0):
    i = _INPUT(type=1)                       # 1 = keyboard
    i.ki = _KEYBDINPUT(vk, scan, flags, 0, 0)
    return i


def send_char(ch):
    """Type one character exactly as it is (works for any character)."""
    data = ch.encode("utf-16-le")
    for k in range(0, len(data), 2):
        unit = int.from_bytes(data[k:k + 2], "little")
        _send([_key(0, unit, KEYEVENTF_UNICODE),
               _key(0, unit, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP)])


def send_vk(vk):
    _send([_key(vk), _key(vk, 0, KEYEVENTF_KEYUP)])


def make_dpi_aware():
    """So pixel positions from Excel match the real screen on scaled displays."""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def bring_to_front():
    try:
        ctypes.windll.user32.ShowWindow(APP.Hwnd, 9)          # restore if minimized
        ctypes.windll.user32.SetForegroundWindow(APP.Hwnd)
    except Exception as e:
        warn_once("front", f"Could not bring Excel to the front: {e}")


def ensure_front():
    """Block until Excel is the window in front, so keystrokes can't go elsewhere."""
    announced, last_try = False, 0.0
    while True:
        try:
            if ctypes.windll.user32.GetForegroundWindow() == APP.Hwnd:
                return
        except Exception:
            return                              # cannot check; carry on
        if not announced:
            log("Excel is not in front - typing paused. Click on Excel to continue.")
            announced = True
        if time.time() - last_try > 20:
            bring_to_front()
            last_try = time.time()
        time.sleep(0.5)


def glide_mouse(x, y):
    """Move the pointer smoothly to (x, y). Movement only, it never clicks."""
    try:
        u = ctypes.windll.user32
        p = _PT()
        u.GetCursorPos(ctypes.byref(p))
        steps = 12
        for i in range(1, steps + 1):
            t = i / steps
            t = t * t * (3 - 2 * t)                           # ease in and out
            u.SetCursorPos(int(p.x + (x - p.x) * t), int(p.y + (y - p.y) * t))
            time.sleep(MOUSE_GLIDE / steps)
    except Exception as e:
        warn_once("mouse", f"Could not move the mouse pointer: {e}")


def show_cell(cell, r):
    """Select the cell about to be typed (the green box) and glide the pointer to it."""
    try:
        win = APP.ActiveWindow
        vis = win.VisibleRange
        if r < vis.Row or r > vis.Row + vis.Rows.Count - 3:
            win.ScrollRow = max(r - 5, 1)                     # scroll the cell into view
        cell.Select()
    except Exception as e:
        warn_once("select", f"Could not select the cell in Excel: {e}")
    if MOVE_MOUSE:
        try:
            win = APP.ActiveWindow
            x = win.PointsToScreenPixelsX(cell.Left + cell.Width / 2)
            y = win.PointsToScreenPixelsY(cell.Top + cell.Height / 2)
            glide_mouse(x, y)
        except Exception as e:
            warn_once("pos", f"Could not find the cell on screen: {e}")


# ------------------------------------------------------------- reading data
def cell_info(c): 
    """Source cell -> (text to type, final value, final number format)."""
    v, fmt = c.value, c.number_format
    if v is None:
        return "", None, fmt
    if isinstance(v, bool):
        return ("TRUE" if v else "FALSE"), v, "General"
    if isinstance(v, str):
        return v, v, ("@" if v.isdigit() else "General")
    if isinstance(v, (int, float)):
        if float(v).is_integer():
            v = int(v)
            show = f"{v:,}" if "#,##0" in fmt else str(v)
        else:
            show = str(v)
        return show, v, fmt
    if isinstance(v, datetime.datetime):
        show = v.strftime("%Y-%m-%d %H:%M:%S") if (v.hour or v.minute or v.second) \
            else v.strftime("%Y-%m-%d")
        return show, v, fmt
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d"), v, fmt
    return str(v), str(v), "General"


def load_source_workbook(path):
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except PermissionError:
        # If the file is currently open in Excel, standard open() fails with
        # ERROR_SHARING_VIOLATION. Read with FILE_SHARE_WRITE so Excel doesn't block it.
        import io
        import msvcrt
        GENERIC_READ = 0x80000000
        FILE_SHARE_READ = 1
        FILE_SHARE_WRITE = 2
        FILE_SHARE_DELETE = 4
        OPEN_EXISTING = 3
        h = ctypes.windll.kernel32.CreateFileW(
            path,
            GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            None,
            OPEN_EXISTING,
            0,
            None,
        )
        if h == -1:
            raise
        try:
            fd = msvcrt.open_osfhandle(h, 0)
            with open(fd, "rb") as f:
                data = f.read()
            return openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        except Exception:
            ctypes.windll.kernel32.CloseHandle(h)
            raise


def read_rows():
    wb = load_source_workbook(SOURCE)   # values, not formulas
    ws = wb[SHEET] if SHEET else wb.worksheets[0]
    rows = [[cell_info(c) for c in r] for r in ws.iter_rows()]
    while rows and not any(i[0] for i in rows[-1]):       # drop trailing blank rows
        rows.pop()
    return rows, ws.title


# ------------------------------------------------------------------ typing
def type_cell(ws, r, c, info, last_in_row):
    show, value, fmt = info
    cell = com(ws.Cells, r, c)
    show_cell(cell, r)                    # cursor box and mouse pointer go to this cell
    if show:
        is_text = isinstance(value, str)
        # text cells are typed into a Text-format cell so Excel never reinterprets them
        com(setattr, cell, "NumberFormat", "@" if is_text else fmt)
        for ch in show:
            ensure_front()
            send_char(ch)                 # edit mode starts, the caret moves with the text
            time.sleep(random.uniform(*CHAR_DELAY))
            if random.random() < HESITATE_CHANCE:
                time.sleep(random.uniform(*HESITATE))
        time.sleep(random.uniform(0.2, 0.6))        # a beat before confirming
    ensure_front()
    send_vk(VK_RETURN if last_in_row else VK_TAB)    # confirm the cell, like a person
    time.sleep(0.25)                                 # let Excel finish committing
    if show and isinstance(value, str) and fmt != "@":
        com(setattr, cell, "NumberFormat", fmt)      # back to the source format
    time.sleep(random.uniform(*CELL_PAUSE))


def type_row(ws, r, row):
    for c, info in enumerate(row, start=1):
        type_cell(ws, r, c, info, last_in_row=(c == len(row)))
    time.sleep(random.uniform(*ROW_PAUSE))


def connect_excel():
    """Start or attach to Excel and wait until it answers."""
    app = None
    try:
        app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        pass
    if app is None:
        try:
            # builds Excel's type information so properties like Visible resolve
            app = win32com.client.gencache.EnsureDispatch("Excel.Application")
        except Exception as e:
            log(f"Note: {e}; trying a dynamic connection")
            try:
                app = win32com.client.Dispatch("Excel.Application")
            except Exception as e2:
                sys.exit(f"Could not start Excel ({e2}). Is Microsoft Excel installed "
                         f"and does it open normally from the Start menu?")
    for _ in range(30):                       # Excel can take a while to start
        try:
            app.Visible = True
            return app
        except Exception:
            time.sleep(2)
    sys.exit("Excel did not respond. Close Excel completely, open it once from the "
             "Start menu, close it, then run the script again. If that fails, "
             "delete the folder %LOCALAPPDATA%\\Temp\\gen_py and retry.")


def open_workbook(app, title, resuming):
    if resuming and os.path.exists(OUTPUT):
        return com(app.Workbooks.Open, OUTPUT)
    wb = com(app.Workbooks.Add)
    com(setattr, wb.Worksheets(1), "Name", title)
    com(wb.SaveAs, OUTPUT, 51)      # 51 = .xlsx
    return wb


def save(wb, done_rows):
    com(wb.Save)
    with open(PROGRESS, "w") as f:
        f.write(str(done_rows))


def main():
    global APP
    keep_awake()
    make_dpi_aware()
    rows, title = read_rows()
    start = 0
    if os.path.exists(PROGRESS):
        start = int(open(PROGRESS).read().strip() or 0)
        if start >= len(rows):
            log("Everything was already typed. Delete the .progress file to redo it.")
            return
        log(f"Resuming at row {start + 1} of {len(rows)}")
    else:
        log(f"{len(rows)} rows to type")

    app = connect_excel()
    APP = app
    app.DisplayAlerts = False
    app.EnableAutoComplete = False       # Excel must not "complete" what is typed
    wb = open_workbook(app, title, start > 0)
    ws = wb.Worksheets(1)
    com(ws.Activate)
    bring_to_front()
    time.sleep(1)

    try:
        begin = last_save = time.time()
        deadline = begin + TARGET_HOURS * 3600
        for n, idx in enumerate(range(start, len(rows)), start=1):
            t0 = time.time()
            type_row(ws, idx + 1, rows[idx])
            if PACED:
                left = max(deadline - time.time(), 0)
                budget = (left + (time.time() - t0)) / (len(rows) - idx)
                factor = random.uniform(0.7, 1.3)
                if random.random() < BREAK_CHANCE:
                    factor *= random.uniform(2.0, 4.0)       # an occasional longer break
                time.sleep(max(min(budget * factor - (time.time() - t0), left), 0))
            if time.time() - last_save > SAVE_MINUTES * 60:
                save(wb, idx + 1)
                last_save = time.time()
            if n % 5 == 0 or idx == len(rows) - 1:
                log(f"row {idx + 1}/{len(rows)} done, "
                    f"{(time.time() - begin) / 3600:.2f} h elapsed")
        com(ws.Columns.AutoFit)
        com(wb.Save)
        if os.path.exists(PROGRESS):
            os.remove(PROGRESS)
        log(f"All rows typed. Saved {OUTPUT}")
    finally:
        try:
            app.EnableAutoComplete = True
        except Exception:
            pass


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log("Stopped. Run again to resume from the last saved row.")
        sys.exit(1)