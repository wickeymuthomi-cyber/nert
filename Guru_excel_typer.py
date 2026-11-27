"""
Guru_excel_typer.py - types the data of an Excel file into Excel at human speed,
character by character, cell by cell, exactly as it appears in the source.

The text is entered with real keystrokes, so each cell goes into edit mode and
you see the blinking text caret moving along as every character appears, just
like typing by hand. The green cell cursor and the mouse pointer move to each
cell first.

Moving the mouse does NOT interrupt it (only the keyboard is used for typing).
If you switch to another window, typing pauses by itself and continues as
soon as the target Excel window is in front again, so no characters are lost
in other programs. Do not type on the keyboard yourself while it runs.

Setup (Windows, Excel installed):
    py -m pip install pywin32 openpyxl
Run:
    py Guru_excel_typer.py
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
import io
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
CELL_PAUSE = (2, 2.5)          # moving on to the next cell
ROW_PAUSE = (3, 3.6)           # finishing a row
COMMIT_SETTLE = 0.25           # fixed settle pause after Tab/Enter

PACED = False                  # True = pace the whole job to TARGET_HOURS
TARGET_HOURS = 7.0             # only used when PACED = True
BREAK_CHANCE = 0.03            # chance of a longer break after a row (PACED)
SAVE_MINUTES = 5
RETRIES = 15                   # attempts per step when Excel is busy
MOVE_MOUSE = True              # the mouse pointer glides to each cell as it is typed
MOUSE_GLIDE = 0.35             # seconds the pointer takes to glide to the next cell

PROGRESS = OUTPUT + ".progress"
APP = None                     # the Excel application, set in main()
TARGET_HWND = None             # cached top-level HWND of the output workbook window
_warned = set()

# Transient Excel COM busy HRESULTs worth retrying
_BUSY_HRESULTS = {
    0x80010001,                # RPC_E_CALL_REJECTED (e.g. cell edit mode)
    0x8001010A,                # RPC_E_SERVERCALL_RETRYLATER
    0x800AC472,                # VBA_E_IGNORE (Excel UI temporarily busy)
}


# ----------------------------------------------------------------- helpers
def log(msg):
    print(time.strftime("[%H:%M:%S] ") + msg, flush=True)


def warn_once(key, msg):
    if key not in _warned:
        _warned.add(key)
        log(msg)


def _is_busy_com_error(exc):
    hr = exc.hresult & 0xFFFFFFFF
    if hr in _BUSY_HRESULTS:
        return True
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo and len(excepinfo) > 5 and excepinfo[5]:
        if (excepinfo[5] & 0xFFFFFFFF) in _BUSY_HRESULTS:
            return True
    return False


def com(fn, *args, retries=None):
    """Run an Excel call, retrying while Excel reports it is busy."""
    max_attempts = RETRIES if retries is None else retries
    for attempt in range(1, max_attempts + 1):
        try:
            return fn(*args)
        except pythoncom.com_error as e:
            if not _is_busy_com_error(e) or attempt == max_attempts:
                raise
            hr = e.hresult & 0xFFFFFFFF
            hint = " (press Esc in Excel if a cell is being edited)" if hr == 0x80010001 else ""
            log(f"Excel busy ({hr:#x}){hint}, retry {attempt}/{max_attempts}")
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
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


KEYEVENTF_KEYUP, KEYEVENTF_UNICODE = 0x2, 0x4
VK_TAB, VK_RETURN, VK_MENU = 0x09, 0x0D, 0x12
SW_RESTORE = 9
GA_ROOT = 2

_win_user32 = None
_win_kernel32 = None


def _user32():
    global _win_user32
    if _win_user32 is None:
        u = ctypes.WinDLL("user32", use_last_error=True)
        u.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(_INPUT), ctypes.c_int]
        u.SendInput.restype = wintypes.UINT
        u.GetForegroundWindow.argtypes = []
        u.GetForegroundWindow.restype = wintypes.HWND
        u.SetForegroundWindow.argtypes = [wintypes.HWND]
        u.SetForegroundWindow.restype = wintypes.BOOL
        u.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
        u.ShowWindow.restype = wintypes.BOOL
        u.IsIconic.argtypes = [wintypes.HWND]
        u.IsIconic.restype = wintypes.BOOL
        u.IsWindow.argtypes = [wintypes.HWND]
        u.IsWindow.restype = wintypes.BOOL
        u.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
        u.GetAncestor.restype = wintypes.HWND
        u.GetCursorPos.argtypes = [ctypes.POINTER(_PT)]
        u.GetCursorPos.restype = wintypes.BOOL
        u.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
        u.SetCursorPos.restype = wintypes.BOOL
        _win_user32 = u
    return _win_user32


def _kernel32():
    global _win_kernel32
    if _win_kernel32 is None:
        k = ctypes.WinDLL("kernel32", use_last_error=True)
        k.SetThreadExecutionState.argtypes = [wintypes.DWORD]
        k.SetThreadExecutionState.restype = wintypes.DWORD
        k.CreateFileW.argtypes = [
            wintypes.LPCWSTR,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.LPVOID,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.HANDLE,
        ]
        k.CreateFileW.restype = wintypes.HANDLE
        k.CloseHandle.argtypes = [wintypes.HANDLE]
        k.CloseHandle.restype = wintypes.BOOL
        _win_kernel32 = k
    return _win_kernel32


def keep_awake():
    """Ask Windows not to sleep while this script runs."""
    try:
        _kernel32().SetThreadExecutionState(0x80000000 | 0x1 | 0x2)
    except Exception:
        log("Could not set keep-awake; disable sleep in Windows power settings.")


def _send(events):
    arr = (_INPUT * len(events))(*events)
    sent = _user32().SendInput(len(events), arr, ctypes.sizeof(_INPUT))
    if sent != len(events):
        err = ctypes.get_last_error()
        raise RuntimeError(
            f"Windows blocked simulated keystrokes (SendInput sent {sent}/{len(events)}, "
            f"error {err}). Make sure Excel and Python run at the same privilege level "
            f"(both normal or both as Administrator)."
        )


def _key(vk=0, scan=0, flags=0):
    i = _INPUT(type=1)                       # 1 = keyboard
    i.ki = _KEYBDINPUT(vk, scan, flags, 0, 0)
    return i


def send_char(ch):
    """Type one character exactly as it is (works for any character)."""
    if ch == "\r":
        return
    if ch == "\n":
        # Alt+Enter inserts an in-cell newline in Excel instead of committing
        _send([
            _key(VK_MENU),
            _key(VK_RETURN),
            _key(VK_RETURN, 0, KEYEVENTF_KEYUP),
            _key(VK_MENU, 0, KEYEVENTF_KEYUP),
        ])
        return
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


def _normalize_hwnd(hwnd):
    """Normalize a COM Long/HWND into a top-level pointer-sized HWND value."""
    if not hwnd:
        return None
    val = ctypes.c_void_p(int(hwnd)).value
    if not val:
        return None
    try:
        root = _user32().GetAncestor(val, GA_ROOT)
        if root:
            return root
    except Exception:
        pass
    return val


def get_target_hwnd(app, wb_win):
    """Resolve and cache the top-level HWND for the target workbook window."""
    raw = None
    if wb_win is not None:
        try:
            raw = com(lambda: wb_win.Hwnd)
        except Exception:
            raw = None
    if not raw:
        raw = com(lambda: app.Hwnd)
    hwnd = _normalize_hwnd(raw)
    if not hwnd:
        raise RuntimeError("Could not determine the Excel window handle (HWND).")
    return hwnd


def is_front():
    """Return True iff the cached target Excel window is the foreground window.

    Uses only Win32 calls (no COM), so it keeps working even while Excel is in
    cell edit mode and rejecting COM requests.
    """
    if not TARGET_HWND:
        return False
    fg = _user32().GetForegroundWindow()
    if not fg:
        return False
    return (int(fg) & 0xFFFFFFFF) == (int(TARGET_HWND) & 0xFFFFFFFF)


def bring_to_front():
    if not TARGET_HWND:
        return
    try:
        u = _user32()
        # Only call SW_RESTORE if minimized so maximized Excel stays maximized
        if u.IsIconic(TARGET_HWND):
            u.ShowWindow(TARGET_HWND, SW_RESTORE)
        u.SetForegroundWindow(TARGET_HWND)
    except Exception as e:
        warn_once("front", f"Could not bring Excel to the front: {e}")


def ensure_front():
    """Block until the target Excel window is in front, so keystrokes can't leak."""
    announced, last_try = False, 0.0
    u = _user32()
    while True:
        if TARGET_HWND and not u.IsWindow(TARGET_HWND):
            raise RuntimeError("The target Excel window was closed.")
        if is_front():
            if announced:
                log("Excel is back in front - resuming typing.")
                time.sleep(0.15)
            return
        if not announced:
            log("Excel is not in front - typing paused. "
                "Click Excel's title bar or taskbar icon to continue.")
            announced = True
        if time.time() - last_try > 20:
            bring_to_front()
            last_try = time.time()
        time.sleep(0.5)


def glide_mouse(x, y, duration=None):
    """Move the pointer smoothly to (x, y). Movement only, it never clicks."""
    try:
        u = _user32()
        p = _PT()
        u.GetCursorPos(ctypes.byref(p))
        steps = 12
        glide_time = MOUSE_GLIDE if duration is None else max(duration, 0.05)
        for i in range(1, steps + 1):
            if not is_front():
                return                              # don't hijack mouse in another app
            t = i / steps
            t = t * t * (3 - 2 * t)                 # ease in and out
            u.SetCursorPos(int(p.x + (x - p.x) * t), int(p.y + (y - p.y) * t))
            time.sleep(glide_time / steps)
    except Exception as e:
        warn_once("mouse", f"Could not move the mouse pointer: {e}")


def show_cell(ws, wb_win, sheet_name, cell, r, speed_scale=1.0):
    """Select the cell about to be typed (the green box) and glide the pointer to it."""
    ensure_front()
    win = wb_win
    if win is None:
        win = com(lambda: APP.ActiveWindow)
    # Keep our worksheet active in case the user switched tabs while paused
    try:
        active_sheet = com(lambda: win.ActiveSheet.Name)
        if active_sheet != sheet_name:
            com(ws.Activate)
    except Exception:
        com(ws.Activate)
    try:
        vis = com(lambda: win.VisibleRange)
        if r < vis.Row or r > vis.Row + vis.Rows.Count - 3:
            com(setattr, win, "ScrollRow", max(r - 5, 1))    # scroll the cell into view
    except Exception as e:
        warn_once("scroll", f"Could not scroll the Excel window: {e}")
    # Selection must succeed; never type into whatever cell happens to be selected
    com(cell.Select)
    if MOVE_MOUSE:
        try:
            left, width, top, height = com(lambda: (cell.Left, cell.Width, cell.Top, cell.Height))
            x = int(round(com(win.PointsToScreenPixelsX, int(round(left + width / 2)))))
            y = int(round(com(win.PointsToScreenPixelsY, int(round(top + height / 2)))))
            glide_mouse(x, y, duration=MOUSE_GLIDE * speed_scale)
        except Exception as e:
            warn_once("pos", f"Could not find the cell on screen: {e}")


# ------------------------------------------------------------- reading data
def cell_info(c, decimal_sep="."):
    """Source cell -> (text to type, final value, final number format)."""
    v = c.value
    fmt = c.number_format or "General"
    if v is None:
        return "", None, fmt
    if isinstance(v, bool):
        return ("TRUE" if v else "FALSE"), v, "General"
    if isinstance(v, str):
        return v, v, ("@" if v.isdigit() else fmt)
    if isinstance(v, (int, float)):
        if float(v).is_integer():
            v = int(v)
            show = str(v)
        else:
            show = str(v)
            if decimal_sep != ".":
                show = show.replace(".", decimal_sep)
        return show, v, fmt
    if isinstance(v, datetime.datetime):
        show = v.strftime("%Y-%m-%d %H:%M:%S") if (v.hour or v.minute or v.second) \
            else v.strftime("%Y-%m-%d")
        return show, v, fmt
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d"), v, fmt
    if isinstance(v, datetime.time):
        return v.strftime("%H:%M:%S"), v, fmt
    return str(v), str(v), "General"


def apply_decimal_separator(rows, decimal_sep):
    """Adapt numeric decimal points in pre-read rows to Excel's active locale."""
    if not decimal_sep or decimal_sep == ".":
        return rows
    adapted = []
    for row in rows:
        new_row = []
        for show, value, fmt in row:
            if isinstance(value, float) and "." in show:
                show = show.replace(".", decimal_sep)
            new_row.append((show, value, fmt))
        adapted.append(new_row)
    return adapted


def load_source_workbook(path):
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except PermissionError:
        # If the file is currently open in Excel, standard open() fails with
        # ERROR_SHARING_VIOLATION. Read with FILE_SHARE_WRITE so Excel doesn't block it.
        import msvcrt

        k32 = _kernel32()
        GENERIC_READ = 0x80000000
        FILE_SHARE_READ = 1
        FILE_SHARE_WRITE = 2
        FILE_SHARE_DELETE = 4
        OPEN_EXISTING = 3
        INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

        h = k32.CreateFileW(
            path,
            GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            None,
            OPEN_EXISTING,
            0,
            None,
        )
        if not h or h == INVALID_HANDLE_VALUE:
            raise ctypes.WinError(ctypes.get_last_error())

        try:
            fd = msvcrt.open_osfhandle(h, os.O_RDONLY)
        except Exception:
            k32.CloseHandle(h)
            raise

        # Ownership of h has transferred to fd; open(fd, ...) takes ownership of fd.
        try:
            f = open(fd, "rb")
        except Exception:
            os.close(fd)
            raise

        with f:
            data = f.read()
        return openpyxl.load_workbook(io.BytesIO(data), data_only=True)


def read_rows():
    wb = load_source_workbook(SOURCE)   # values, not formulas
    try:
        ws = wb[SHEET] if SHEET else wb.worksheets[0]
        rows = [[cell_info(c) for c in r] for r in ws.iter_rows()]
        while rows and not any(i[0] for i in rows[-1]):       # drop trailing blank rows
            rows.pop()
        return rows, ws.title
    finally:
        wb.close()


def estimate_typing_seconds(rows):
    """Return (scalable_seconds, fixed_seconds) estimated to type rows unpaced."""
    avg_char = sum(CHAR_DELAY) / 2 + HESITATE_CHANCE * (sum(HESITATE) / 2)
    avg_cell_scalable = (MOUSE_GLIDE if MOVE_MOUSE else 0.0) + 0.4 + sum(CELL_PAUSE) / 2
    avg_row_scalable = sum(ROW_PAUSE) / 2
    scalable = 0.0
    fixed = 0.0
    for row in rows:
        scalable += avg_row_scalable
        for show, _, _ in row:
            scalable += avg_cell_scalable
            fixed += COMMIT_SETTLE
            if show:
                scalable += len(show) * avg_char
    return scalable, fixed


# ------------------------------------------------------------------ typing
def type_cell(ws, wb_win, sheet_name, r, c, info, last_in_row, speed_scale=1.0):
    show, value, fmt = info
    cell = com(ws.Cells, r, c)
    show_cell(ws, wb_win, sheet_name, cell, r, speed_scale=speed_scale)
    if show:
        is_text = isinstance(value, str)
        # text cells are typed into a Text-format cell so Excel never reinterprets them
        com(setattr, cell, "NumberFormat", "@" if is_text else fmt)
        for ch in show:
            ensure_front()
            send_char(ch)                 # edit mode starts, the caret moves with the text
            time.sleep(random.uniform(*CHAR_DELAY) * speed_scale)
            if random.random() < HESITATE_CHANCE:
                time.sleep(random.uniform(*HESITATE) * speed_scale)
        time.sleep(random.uniform(0.2, 0.6) * speed_scale)    # a beat before confirming
    ensure_front()
    send_vk(VK_RETURN if last_in_row else VK_TAB)    # confirm the cell, like a person
    time.sleep(COMMIT_SETTLE)                        # let Excel finish committing
    if show and isinstance(value, str) and fmt != "@":
        com(setattr, cell, "NumberFormat", fmt)      # back to the source format
    time.sleep(random.uniform(*CELL_PAUSE) * speed_scale)


def type_row(ws, wb_win, sheet_name, r, row, speed_scale=1.0):
    for c, info in enumerate(row, start=1):
        type_cell(
            ws,
            wb_win,
            sheet_name,
            r,
            c,
            info,
            last_in_row=(c == len(row)),
            speed_scale=speed_scale,
        )
    time.sleep(random.uniform(*ROW_PAUSE) * speed_scale)


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


def get_excel_decimal_separator(app):
    """Query Excel's active decimal separator (xlDecimalSeparator = 3)."""
    try:
        sep = com(app.International, 3)
        if isinstance(sep, str) and len(sep) == 1:
            return sep
    except Exception:
        pass
    return "."


def find_open_workbook(app, path):
    """Find an already-open workbook matching path (or conflicting by basename)."""
    target_path = os.path.normcase(os.path.abspath(path))
    target_name = os.path.basename(target_path)
    exact_match = None
    name_conflict = None
    try:
        for wb in app.Workbooks:
            try:
                wb_full = os.path.normcase(os.path.abspath(wb.FullName))
                wb_name = os.path.normcase(wb.Name)
            except Exception:
                continue
            if wb_full == target_path:
                exact_match = wb
                break
            if wb_name == target_name:
                name_conflict = wb
    except Exception:
        pass
    return exact_match, name_conflict


def open_workbook(app, title, resuming):
    exact, conflict = find_open_workbook(app, OUTPUT)
    if conflict is not None:
        sys.exit(
            f"Another workbook named '{os.path.basename(OUTPUT)}' is open in Excel "
            f"from a different folder. Close it first, then run the script again."
        )
    if exact is not None:
        # Close any previously open copy without saving so resuming reloads the exact
        # last-saved checkpoint and fresh starts don't fail on SaveAs.
        com(exact.Close, False)

    if resuming and os.path.exists(OUTPUT):
        return com(app.Workbooks.Open, os.path.abspath(OUTPUT))

    wb = com(app.Workbooks.Add)
    ws1 = com(lambda: wb.Worksheets(1))
    com(setattr, ws1, "Name", title)
    prev_alerts = com(lambda: app.DisplayAlerts)
    try:
        com(setattr, app, "DisplayAlerts", False)
        com(wb.SaveAs, os.path.abspath(OUTPUT), 51)      # 51 = .xlsx
    finally:
        try:
            com(setattr, app, "DisplayAlerts", prev_alerts, retries=3)
        except Exception:
            pass
    return wb


def load_progress():
    if not os.path.exists(PROGRESS):
        return 0
    with open(PROGRESS, encoding="utf-8") as f:
        raw = f.read().strip()
    if not raw:
        return 0
    try:
        value = int(raw)
    except ValueError:
        sys.exit(
            f"Corrupted progress file '{PROGRESS}' (contains {raw!r}). "
            f"Fix or delete it to start over."
        )
    if value < 0:
        sys.exit(
            f"Invalid progress file '{PROGRESS}' (negative row {value}). "
            f"Delete it to start over."
        )
    return value


def save(wb, done_rows):
    com(wb.Save)
    tmp = PROGRESS + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(str(done_rows))
    os.replace(tmp, PROGRESS)


def main():
    global APP, TARGET_HWND
    keep_awake()
    make_dpi_aware()
    rows, title = read_rows()
    if not rows:
        log("Source sheet has no data rows to type.")
        return

    start = load_progress()
    if start >= len(rows):
        log("Everything was already typed. Delete the .progress file to redo it.")
        return
    if start > 0 and not os.path.exists(OUTPUT):
        log(f"Progress file points to row {start + 1}, but {OUTPUT} is missing. "
            f"Starting over from row 1.")
        start = 0

    scalable_s, fixed_s = estimate_typing_seconds(rows[start:])
    est_total_s = scalable_s + fixed_s
    speed_scale = 1.0
    target_seconds = 0.0
    if PACED:
        target_seconds = TARGET_HOURS * 3600 * (len(rows) - start) / len(rows)
        budget_scalable = target_seconds * 0.85 - fixed_s
        if est_total_s > target_seconds * 0.85 and scalable_s > 0:
            if budget_scalable <= 0:
                speed_scale = 0.02
                log(f"Warning: TARGET_HOURS ({target_seconds / 3600:.2f} h for remaining rows) "
                    f"is shorter than minimum safe commit time (~{fixed_s / 3600:.2f} h). "
                    f"Typing as fast as safely possible.")
            else:
                speed_scale = max(budget_scalable / scalable_s, 0.02)
                log(f"PACED mode: unpaced run (~{est_total_s / 3600:.2f} h) exceeds target "
                    f"({target_seconds / 3600:.2f} h); scaling typing delays by {speed_scale:.2f}x")

    if start > 0:
        log(f"Resuming at row {start + 1} of {len(rows)} "
            f"(est. ~{(scalable_s * speed_scale + fixed_s) / 3600:.2f} h remaining)")
    else:
        log(f"{len(rows)} rows to type "
            f"(est. ~{(scalable_s * speed_scale + fixed_s) / 3600:.2f} h)")

    app = connect_excel()
    APP = app
    decimal_sep = get_excel_decimal_separator(app)
    rows = apply_decimal_separator(rows, decimal_sep)

    orig_display_alerts = com(lambda: app.DisplayAlerts)
    orig_autocomplete = com(lambda: app.EnableAutoComplete)
    com(setattr, app, "EnableAutoComplete", False)       # Excel must not "complete" what is typed
    try:
        wb = open_workbook(app, title, start > 0)
        ws = com(lambda: wb.Worksheets(1))
        com(wb.Activate)
        com(ws.Activate)
        try:
            wb_win = com(lambda: wb.Windows(1))
        except Exception:
            wb_win = None
        TARGET_HWND = get_target_hwnd(app, wb_win)
        bring_to_front()
        time.sleep(1)

        begin = last_save = time.time()
        deadline = begin + target_seconds
        for n, idx in enumerate(range(start, len(rows)), start=1):
            t0 = time.time()
            type_row(ws, wb_win, title, idx + 1, rows[idx], speed_scale=speed_scale)
            remaining_rows = len(rows) - idx - 1
            if PACED and remaining_rows > 0:
                now = time.time()
                row_time = now - t0
                left = max(deadline - now, 0.0)
                budget = (left + row_time) / (remaining_rows + 1)
                factor = random.uniform(0.7, 1.3)
                if random.random() < BREAK_CHANCE:
                    factor *= random.uniform(2.0, 4.0)       # an occasional longer break
                time.sleep(max(min(budget * factor - row_time, left), 0.0))
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
        for prop, val in (("EnableAutoComplete", orig_autocomplete),
                          ("DisplayAlerts", orig_display_alerts)):
            try:
                com(setattr, app, prop, val, retries=2)
            except Exception:
                log(f"Note: could not restore Excel.{prop} (Excel may still be in edit mode).")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log("Stopped. Run again to resume from the last saved row.")
        sys.exit(1)
    except RuntimeError as e:
        log(f"Stopped: {e}")
        sys.exit(1)
