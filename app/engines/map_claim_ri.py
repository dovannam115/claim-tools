# -*- coding: utf-8 -*-
"""
CLAIM -> POLICY / RI MAPPING TOOL  (Equipment - Claims dept)
===========================================================
Monthly: put the claim file (e.g. "DOT XX_..._EW_ Tmm.yyyy.xlsx") and double-click
"MAP CLAIM.bat" (pick the file in the dialog), or run:

    python map_claim_ri.py                  # opens a file picker
    python map_claim_ri.py "path\\file.xlsx" # use a specific file

OUTPUT (ONE file): <name>_MAPPED.xlsx  (keeps the original format/header/logo; appends new
columns after the last column on the "140 HS" sheet (from STT=1) AND the "Data" sheet):
    Policy No | Program | Program Coverage | Effective Date | Expired Date | SI | RI Share Pct
    + Hannover-re | Hanoi-re | TMIV | Fubon   (share tung nha tai theo hop dong tai)
    + Luu y HSBT (claim review note, kept in Vietnamese)
"RI Share Pct" = TONG share cac nha tai theo hop dong tai (bang dbo.D_RI_STRUCTURE),
KHONG phai cot [RI Share Pct] cua bang phi trong database.
The PIAS input form is added as an extra sheet named "Pias" INSIDE the same file
(no separate _Pias.xlsx), built from "Template nhap HSBT Pias.xlsx" in this folder.

MATCH KEY (priority): Serial of product = premium[IMEI]; fallback Mobile No.(last 9)=premium[Phone].
When one serial hits several policies: prefer SI = claim SI -> nearest effective date -> named program -> base.
DB: Fabric Warehouse 'DataWarehouse - Equipment', view dbo.vw_Claim_RI_Lookup (exposes 7 fields + keys).
"""
import os, sys, io, glob, struct, datetime, re, zipfile, subprocess, tempfile
# Host kieu Streamlit/pythonw co the khong co stdout that, hoac dong no giua chung.
# KHONG boc them mot TextIOWrapper moi quanh sys.stdout.buffer: ban cu mat tham chieu,
# bi thu don, va khi bi thu don no DONG luon buffer ben duoi -> moi print sau do nem
# "ValueError: I/O operation on closed file" (da lam gay Map Claim tren web 22/08/2026).
try:
    sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
except Exception:
    pass

_print = print


def print(*a, **k):        # noqa: A001 - engine von chay o console, giu nguyen moi loi goi print
    """In khi con console. Chay duoi web thi im lang chu khong lam vo ca lan chay."""
    try:
        _print(*a, **k)
    except (ValueError, OSError):
        pass


try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
except ImportError as _e:
    print("\n*** ERROR: Missing Python libraries (" + str(_e) + ").")
    print("*** Please launch the tool via 'MAP CLAIM.bat' (it uses the bundled Python).\n")
    sys.exit(1)

HERE   = os.path.dirname(os.path.abspath(__file__))
AUTH   = os.path.join(HERE, "fabric_auth_record.json")   # auth record canh script -> portable
SERVER = "ondmtv5n6agudnvvpb5ctkglh4-remsb7fj4tuebamhm7fnutbblm.datawarehouse.fabric.microsoft.com"
DBNAME = "DataWarehouse - Equipment"
CLIENT = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"          # Azure CLI public client (device-code)
VIEWNM = "dbo.vw_Claim_RI_Lookup"

# ---------- GUI (Windows dialogs via PowerShell/WinForms; no extra install) ----------
GUI = False                    # set True at startup if PowerShell works -> use boxes, hide console

# Set True by a server host (app.py) that imports this module as a library. There is
# nobody at the console to read a message or press a key, so die() must raise instead
# of printing + waiting on input() -- which would hang the request forever.
HEADLESS = False
class MapError(Exception):
    """A user-facing failure raised by die() when HEADLESS."""
_PS_DIALOG = r'''
param([string]$Mode,[string]$Title,[string]$Prompt,[string]$ItemsFile,[string]$Icon="Information")
Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
if ($Mode -eq "msg") {
    $ic = [System.Windows.Forms.MessageBoxIcon]::$Icon
    [void][System.Windows.Forms.MessageBox]::Show($Prompt,$Title,[System.Windows.Forms.MessageBoxButtons]::OK,$ic)
    exit 0
}
if ($Mode -eq "progress") {
    $f = New-Object System.Windows.Forms.Form
    $f.Text="MAP CLAIM"; $f.StartPosition="CenterScreen"; $f.Size=New-Object System.Drawing.Size(370,120)
    $f.FormBorderStyle="FixedToolWindow"; $f.TopMost=$true; $f.ControlBox=$false
    $l = New-Object System.Windows.Forms.Label
    $l.Text=$Prompt; $l.TextAlign="MiddleCenter"; $l.Dock="Fill"; $l.Font=New-Object System.Drawing.Font("Segoe UI",11)
    $f.Controls.Add($l); $f.Add_Shown({$f.Activate()})
    [void]$f.ShowDialog(); exit 0
}
$items = @(Get-Content -LiteralPath $ItemsFile -Encoding UTF8)
$f = New-Object System.Windows.Forms.Form
$f.Text=$Title; $f.StartPosition="CenterScreen"; $f.Size=New-Object System.Drawing.Size(470,200)
$f.FormBorderStyle="FixedDialog"; $f.MaximizeBox=$false; $f.MinimizeBox=$false; $f.TopMost=$true
$l = New-Object System.Windows.Forms.Label
$l.Text=$Prompt; $l.Size=New-Object System.Drawing.Size(440,42); $l.Location=New-Object System.Drawing.Point(15,12)
$f.Controls.Add($l)
$c = New-Object System.Windows.Forms.ComboBox
$c.DropDownStyle="DropDownList"; $c.Size=New-Object System.Drawing.Size(430,28); $c.Location=New-Object System.Drawing.Point(15,60)
foreach ($it in $items) { [void]$c.Items.Add($it) }
if ($c.Items.Count -gt 0) { $c.SelectedIndex=0 }
$f.Controls.Add($c)
$ok = New-Object System.Windows.Forms.Button
$ok.Text="OK"; $ok.Size=New-Object System.Drawing.Size(95,32); $ok.Location=New-Object System.Drawing.Point(250,110); $ok.DialogResult="OK"
$f.Controls.Add($ok); $f.AcceptButton=$ok
$cx = New-Object System.Windows.Forms.Button
$cx.Text="Cancel"; $cx.Size=New-Object System.Drawing.Size(95,32); $cx.Location=New-Object System.Drawing.Point(352,110); $cx.DialogResult="Cancel"
$f.Controls.Add($cx); $f.CancelButton=$cx
$f.Add_Shown({$f.Activate()})
if ($f.ShowDialog() -eq "OK") { [Console]::Out.WriteLine($c.SelectedIndex) } else { [Console]::Out.WriteLine(-1) }
'''
def _ps_script_path():
    p = os.path.join(tempfile.gettempdir(), "claimri_dialog.ps1")
    try:
        if not os.path.exists(p) or os.path.getsize(p) != len(_PS_DIALOG.encode("utf-8")):
            open(p, "w", encoding="utf-8").write(_PS_DIALOG)
    except Exception:
        return None
    return p
_NO_WINDOW = 0x08000000            # CREATE_NO_WINDOW: don't flash a console for the child powershell
def _ps_run(args, timeout=600):
    ps = _ps_script_path()
    if not ps: return None
    try:
        r = subprocess.run(["powershell","-STA","-NoProfile","-ExecutionPolicy","Bypass","-File",ps] + args,
                           capture_output=True, text=True, encoding="utf-8", timeout=timeout,
                           creationflags=_NO_WINDOW)
        return (r.stdout or "").strip()
    except Exception:
        return None
def ui_choice(title, prompt, options):
    """Dropdown dialog -> selected index (int) or None (cancel / unavailable)."""
    itf = os.path.join(tempfile.gettempdir(), "claimri_items.txt")
    try: open(itf, "w", encoding="utf-8").write("\n".join(options))
    except Exception: return None
    out = _ps_run(["-Mode","choice","-Title",title,"-Prompt",prompt,"-ItemsFile",itf])
    if out is None: return None
    try:
        i = int(out.splitlines()[-1]); return i if i >= 0 else None
    except Exception:
        return None
def ui_msg(title, message, icon="Information"):
    _ps_run(["-Mode","msg","-Title",title,"-Prompt",message,"-Icon",icon], timeout=1800)
def ui_progress_start(text):
    """Non-blocking 'please wait' window; returns a handle to stop later."""
    if not GUI: return None
    ps = _ps_script_path()
    if not ps: return None
    try:
        return subprocess.Popen(["powershell","-STA","-NoProfile","-ExecutionPolicy","Bypass","-File",ps,
                                 "-Mode","progress","-Prompt",text], creationflags=_NO_WINDOW)
    except Exception:
        return None
def ui_progress_stop(p):
    try:
        if p: p.terminate()
    except Exception: pass
def ui_signin(url, code):
    try: os.startfile(url)
    except Exception:
        try:
            import webbrowser; webbrowser.open(url)
        except Exception: pass
    ui_msg("Fabric sign-in (first run)",
           "A browser was opened to sign in to Fabric.\n\nIf not, open this link:\n" + url +
           "\n\nEnter this code:   " + code + "\n\nSign in with a PVI account that can access the "
           "Equipment workspace, then click OK.", "Information")
def _have_powershell():
    try:
        r = subprocess.run(["powershell","-NoProfile","-Command","exit 7"], capture_output=True, timeout=20,
                           creationflags=0x08000000)
        return r.returncode == 7
    except Exception:
        return False
def _hide_console():
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd: ctypes.windll.user32.ShowWindow(hwnd, 0)      # SW_HIDE
    except Exception:
        pass
def _unhide_console():
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd: ctypes.windll.user32.ShowWindow(hwnd, 5)      # SW_SHOW
    except Exception:
        pass

def die(msg):
    if HEADLESS:
        raise MapError(msg)
    if GUI:
        ui_msg("MAP CLAIM - Error", msg + "\n\n(Screenshot this and send to Nam if you need help.)", "Error")
    print("\n*** ERROR: " + msg); print("*** (screenshot this and send to Nam if you need help)\n")
    if not GUI:
        try: input("\nPress Enter to close...")
        except Exception: pass
    sys.exit(1)

# ---------- helpers ----------
def norm(v): return None if v is None else str(v).strip()
def stext(v):
    """Text hoa 1 o DINH DANH (IMEI/Serial, so yeu cau, SDT, ma giao dich, so don).
       Excel hay luu IMEI toan chu so thanh SO -> str() se ra '3.51955e+14' hoac duoi '.0',
       va neu ghi thang ra file thi Excel hien E+14. Luon tra ve chuoi chu so nguyen.
       DUNG HAM NAY cho MOI cho co IMEI/dinh danh - doc, ghi, lan preview web."""
    if v is None: return None
    if isinstance(v, bool): return str(v)
    if isinstance(v, int): return str(v)
    if isinstance(v, float): return str(int(v)) if v.is_integer() else repr(v)
    return str(v).strip()
def last9(v):
    s = stext(v)                                   # stext: SDT dang so khong bi dinh '.0' -> lech 1 chu so
    if s is None: return None
    s = "".join(ch for ch in s if ch.isdigit())
    return s[-9:] if len(s) >= 9 else None
def pdate(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date):     return v
    s = norm(v)
    if not s: return None
    s = s.split()[0]
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
        try: return datetime.datetime.strptime(s, f).date()
        except: pass
    return None
def pnum(v):
    if v is None: return None
    try: return float(str(v).replace(",", "").strip())
    except: return None
_XL_EPOCH = datetime.date(1899, 12, 30)          # Excel serial-date base (1900 leap-bug offset)
def xldate(v):
    d = pdate(v)
    return (d - _XL_EPOCH).days if d else None

# ---------- 1. locate input ----------
def choose_file_dialog(initial_dir, title):
    """Hop thoai Open-File cua Windows qua ctypes (chay duoc trong Python nhung, khong can tkinter).
       Tra ve: duong dan (chon) / None (bam Cancel) / False (khong mo duoc hop thoai)."""
    try:
        import ctypes
        from ctypes import wintypes
        class OPENFILENAME(ctypes.Structure):
            _fields_ = [
                ("lStructSize", wintypes.DWORD), ("hwndOwner", wintypes.HWND),
                ("hInstance", wintypes.HINSTANCE), ("lpstrFilter", wintypes.LPCWSTR),
                ("lpstrCustomFilter", wintypes.LPWSTR), ("nMaxCustFilter", wintypes.DWORD),
                ("nFilterIndex", wintypes.DWORD), ("lpstrFile", wintypes.LPWSTR),
                ("nMaxFile", wintypes.DWORD), ("lpstrFileTitle", wintypes.LPWSTR),
                ("nMaxFileTitle", wintypes.DWORD), ("lpstrInitialDir", wintypes.LPCWSTR),
                ("lpstrTitle", wintypes.LPCWSTR), ("Flags", wintypes.DWORD),
                ("nFileOffset", wintypes.WORD), ("nFileExtension", wintypes.WORD),
                ("lpstrDefExt", wintypes.LPCWSTR), ("lCustData", wintypes.LPARAM),
                ("lpfnHook", wintypes.LPVOID), ("lpTemplateName", wintypes.LPCWSTR),
                ("pvReserved", wintypes.LPVOID), ("dwReserved", wintypes.DWORD),
                ("FlagsEx", wintypes.DWORD)]
        buf = ctypes.create_unicode_buffer(2048)
        ofn = OPENFILENAME()
        ofn.lStructSize = ctypes.sizeof(ofn)
        ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
        ofn.nMaxFile = 2048
        ofn.lpstrFilter = "Excel (*.xlsx;*.xlsm)\0*.xlsx;*.xlsm\0All files (*.*)\0*.*\0"
        ofn.nFilterIndex = 1
        ofn.lpstrInitialDir = initial_dir
        ofn.lpstrTitle = title
        ofn.Flags = 0x00081800  # EXPLORER | FILEMUSTEXIST | PATHMUSTEXIST
        if ctypes.windll.comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
            return buf.value
        return None
    except Exception:
        return False

def pick_input():
    if len(sys.argv) > 1:                       # keo-tha file vao .bat / chi dinh duong dan
        p = sys.argv[1]
        if not os.path.isabs(p): p = os.path.join(HERE, p)
        if not os.path.exists(p): die(f"Khong thay file: {p}")
        return p
    print("  Opening file picker ...")
    sel = choose_file_dialog(HERE, "Select the claim file (.xlsx) to map")
    if sel:
        return sel
    if sel is None:
        die("No file selected. Run again and pick the claim .xlsx (or drag-drop the file onto MAP CLAIM.bat).")
    # sel is False -> dialog unavailable -> fall back to the newest .xlsx in this folder
    cands = [f for f in glob.glob(os.path.join(HERE, "*.xlsx"))
             if not os.path.basename(f).startswith("~$")
             and not f.upper().endswith("_MAPPED.XLSX")]
    if not cands:
        die("File picker unavailable and no .xlsx found in this folder.")
    cands.sort(key=os.path.getmtime, reverse=True)
    print("  (picker unavailable -> using newest file:", os.path.basename(cands[0]), ")")
    return cands[0]

# ---------- 2. read claim 'Data' sheet (by header name) ----------
ALIASES = {
    "rr":     ["Require Repair No", "Claim Request №", "Claim Request No", "Reception №"],
    "serial": ["Serial of product", "IMEI/Serial", "Device IMEI", "Serial"],
    "phone":  ["Mobile No.", "Mobile No", "Phone", "SĐT"],
    "mgd":    ["Mã giao dịch", "Transaction Code", "RequestID", "Ma giao dich"],
    "no":     ["No", "№", "STT"],
    "si":     ["Sum Insured", "SI"],
    "incident":["Incident Date", "Dol Date of Loss", "Date of Loss"],
    "notif":  ["Notification Date", "Notified Date"],
    "cause":  ["Cause of damage", "Cause of Loss", "Cause of loss"],
    "detail": ["Detail of damage", "Detail of Loss"],
    "cost":   ["Total indemnity Amount", "Approved Amount", "Claim Amount"],
    "eff":    ["Ngày hiệu lực", "Effective Date"],
    "exp":    ["Ngày hết hiệu lực", "Expired Date"],
}
def find_header_row(ws, want):
    want = [w.lower() for w in want]
    best = (0, None, None)
    for ri in range(1, min(ws.max_row, 15) + 1):
        vals = {}
        n = 0
        for c in ws[ri]:
            if c.value is not None:
                vals[norm(c.value)] = c.column - 1
                if norm(c.value).lower() in want: n += 1
        if n > best[0]: best = (n, ri, vals)
    return (best[1], best[2]) if best[0] >= 2 else (None, None)

def _hdr_row_cells(ws, r):
    row = next(ws.iter_rows(min_row=r, max_row=r), None)
    return row or ()
def _bangke_header_row(ws):
    """1-based header row of a 'bang ke' sheet ('Số yêu cầu' + 'IMEI/Serial'), else None."""
    for r in range(1, (min(ws.max_row, 20) if ws.max_row else 20) + 1):
        vals = {norm(c.value) for c in _hdr_row_cells(ws, r) if c.value is not None}
        if "Số yêu cầu" in vals and "IMEI/Serial" in vals:
            return r
    return None
def _hdr_index(ws, hr):
    return {norm(c.value): c.column - 1 for c in _hdr_row_cells(ws, hr) if c.value is not None}
def _row_key_cols(ws, hr, idx):
    """(cot STT, cot IMEI dung lam du phong hoac None).

    Binh thuong dong du lieu = dong co STT la so khac 0; bo loc nay chinh la thu chan
    cac dong tong cong / ghi chu nam duoi bang. Nhung co bang ke xuat tu he thong khac
    BO TRONG han cot STT (vd 'Total Claim Q2.2026_Check HSBT' sheet 'Mapped') -> khong con
    dong nao doc duoc. Chi trong truong hop cot STT rong/thieu TOAN sheet moi chuyen sang
    nhan dong bang cot IMEI/Serial (dong tong cong khong co IMEI nen van bi loai)."""
    stt_c = idx.get("STT")
    if stt_c is not None:
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            v = row[stt_c] if stt_c < len(row) else None
            if isinstance(v, (int, float)) and v:
                return stt_c, None
    for a in BK_AL["serial"]:
        if a in idx: return stt_c, idx[a]
    return stt_c, None
def _is_data_row(vals, stt_c, ser_c):
    v = vals[stt_c] if (stt_c is not None and stt_c < len(vals)) else None
    if isinstance(v, (int, float)) and v: return True
    if ser_c is None: return False
    s = vals[ser_c] if ser_c < len(vals) else None
    return s is not None and str(s).strip() != ""
def _count_stt_rows(ws, hr, idx):
    stt_c, ser_c = _row_key_cols(ws, hr, idx)
    return sum(1 for row in ws.iter_rows(min_row=hr + 1, values_only=True)
               if _is_data_row(row, stt_c, ser_c))

# ---------- 2b. list mappable sheets + let the user pick one ----------
def list_mappable(path):
    """[dict(sheet, kind, hr, n)] for every sheet we can map (Vietnamese bang ke or English 'Data')."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        die(f"Cannot open the input Excel file. It may be open in Excel -> close it. ({e})")
    out = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        hr = _bangke_header_row(ws)
        if hr:
            idx = _hdr_index(ws, hr)
            out.append(dict(sheet=sn, kind="bangke", hr=hr, n=_count_stt_rows(ws, hr, idx)))
            continue
        hr2, idx = find_header_row(ws, ALIASES["serial"] + ALIASES["rr"])
        if idx and any(a in idx for a in ALIASES["serial"] + ALIASES["rr"]):
            out.append(dict(sheet=sn, kind="data", hr=hr2, n=0))
    wb.close()
    return out

def choose_mapping(path):
    cands = list_mappable(path)
    if not cands:
        die("No mappable sheet found (need a Vietnamese bang ke with 'Số yêu cầu' + 'IMEI/Serial', "
            "or an English 'Data' sheet with 'Serial of product' + 'Require Repair No').")
    if len(sys.argv) > 2:                       # explicit sheet: python map_claim_ri.py file.xlsx "Sheet name"
        want = norm(sys.argv[2])
        for c in cands:
            if c["sheet"] == want:
                print(f"  Sheet (from argument): '{c['sheet']}'"); return c
        die("Sheet '%s' is not mappable. Mappable sheets: %s"
            % (want, ", ".join(repr(c["sheet"]) for c in cands)))
    if len(cands) == 1:
        print(f"  Only one mappable sheet -> '{cands[0]['sheet']}'"); return cands[0]
    labels = [f"{c['sheet']}  ({'bang ke, %d rows' % c['n'] if c['kind']=='bangke' else 'Data (English)'})"
              for c in cands]
    if GUI:
        i = ui_choice("MAP CLAIM - Choose sheet", "This file has several sheets. Which one to map?", labels)
        if i is None: die("No sheet selected.")
        print(f"  Sheet (dialog): '{cands[i]['sheet']}'"); return cands[i]
    print("\n  This file has several sheets. Which one do you want to map?")
    for i, lb in enumerate(labels, 1):
        print(f"    [{i}] {lb}")
    while True:
        try:
            ans = input("\n  Enter the number of the sheet to map: ").strip()
        except EOFError:
            die("No sheet selected. Re-run and pick a sheet, or pass it as the 2nd argument.")
        if ans.isdigit() and 1 <= int(ans) <= len(cands):
            return cands[int(ans) - 1]
        print("  Invalid choice, try again.")

# (label, ds-group, product-family-code). Each batch is ONE product; the family code
# mirrors dbo.vw_Claim_Premium_Matched (020413 / 020415 / 020407 / 020408).
PRODUCTS = [
    ("1 doi 1  (1D1 / MDM)",           "1D1", "020413"),
    ("Nut man hinh  (Crack / NVMH)",   "MT",  "020415"),
    ("Roi vo / Tai nan  (ADLD)",       "MT",  "020407"),
    ("Bao hanh mo rong  (EW / ICT)",   "MT",  "020408"),
    ("All products (do not filter)",   None,  None),
]
_PROD_ALIAS = {"1D1":0,"MDM":0,"020413":0, "CRACK":1,"NVMH":1,"020415":1,
    "ADLD":2,"AD":2,"020407":2, "EW":3,"ICT":3,"020408":3, "ALL":4,"*":4}
def choose_product():
    """Which insurance PRODUCT is THIS batch? -> (ds_group, family_code, label).
       Each claim batch is one product; picking it (a) hard-filters to that ds-group
       (1D1 vs MT) and (b) prefers the matching product family while ranking, so a
       'Crack' batch grabs the Crack policy, not the EW policy of the same device.
       Headless: 4th CLI arg = 1D1 / CRACK / ADLD / EW / ALL (or the 020xxx code)."""
    idx = None
    if len(sys.argv) > 3:
        a = norm(sys.argv[3]).upper()
        if a.isdigit() and 1 <= int(a) <= len(PRODUCTS): idx = int(a) - 1
        elif a in _PROD_ALIAS: idx = _PROD_ALIAS[a]
        if idx is not None:
            lbl, g, f = PRODUCTS[idx]; print(f"  Product (from argument): {lbl}"); return g, f, lbl
    if GUI:
        i = ui_choice("MAP CLAIM - Choose product", "Which insurance PRODUCT is this batch?",
                      [p[0] for p in PRODUCTS])
        if i is None: die("No product selected.")
        lbl, g, f = PRODUCTS[i]; print(f"  Product (dialog): {lbl}"); return g, f, lbl
    print("\n  Which insurance PRODUCT is THIS batch?")
    for i, (lbl, g, f) in enumerate(PRODUCTS, 1):
        print(f"    [{i}] {lbl}")
    while True:
        try:
            ans = input("\n  Enter number [default 1]: ").strip() or "1"
        except EOFError:
            lbl, g, f = PRODUCTS[0]; print(f"  (no input -> {lbl})"); return g, f, lbl
        if ans.isdigit() and 1 <= int(ans) <= len(PRODUCTS):
            lbl, g, f = PRODUCTS[int(ans) - 1]; print(f"  Product: {lbl}"); return g, f, lbl
        print("  Invalid choice, try again.")

# ---------- 2c. read the chosen sheet ----------
def read_data(path, sheet):
    """Read an English 'Data' sheet -> (claims, header_row)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    hr, idx = find_header_row(ws, ALIASES["serial"] + ALIASES["rr"])
    def col(key):
        for a in ALIASES[key]:
            if a in idx: return idx[a]
        return None
    cmap = {k: col(k) for k in ALIASES}
    claims = []
    for row in ws.iter_rows(min_row=hr + 1, values_only=False):
        vals = [c.value for c in row]
        if all(v is None for v in vals): continue
        def g(k):
            i = cmap[k]; return vals[i] if (i is not None and i < len(vals)) else None
        if g("serial") is None and g("rr") is None and g("phone") is None: continue
        claims.append(dict(sheet_row=row[0].row, no=g("no"), rr=stext(g("rr")),
            serial=stext(g("serial")), phone=stext(g("phone")), phone9=last9(g("phone")),
            mgd=stext(g("mgd")), si=pnum(g("si")), incident=pdate(g("incident")), notif=pdate(g("notif")),
            cause=norm(g("cause")), detail=norm(g("detail")), cost=pnum(g("cost")),
            eff_file=pdate(g("eff")), exp_file=pdate(g("exp"))))
    wb.close()
    return claims, hr

BK_AL = {
    "stt":["STT"], "gsc":["Số yêu cầu"], "serial":["IMEI/Serial"],
    "phone":["Số điện thoại khách hàng", "Số điện thoại", "SĐT", "Số ĐT", "Phone"],
    "mgd":["Mã giao dịch", "Ma giao dich", "Mã GD"],
    "si":["Số tiền bảo hiểm"], "incident":["Ngày tổn thất"], "notif":["Ngày thông báo"],
    "cause":["Nguyên nhân hư hỏng", "Mô tả nguyên nhân hư hỏng"],
    "eff":["Ngày hiệu lực"], "exp":["Ngày hết hiệu lực"],
    "prod":["Tên sản phẩm"],
}
# Chi phi BT = cot "Phi tra GSC (+vat)", do theo TEN header - KHONG lay cot cuoi cung.
# (Bang ke doi moi con them cot duoi "VAT linh kien"/"So HD bao hiem"/"So dien thoai";
#  lay cot cuoi se doc nham SDT thanh so tien -> moi dong deu bao "Chi phi BT > 100% SI".)
COST_HDRS = ["Phí trả GSC (+vat)", "Tổng phí trả GSC (+vat)", "Tổng phí (+vat)"]
def _sq(s):
    """Gop moi khoang trang (ke ca xuong dong trong header) + bo dau cach + lower."""
    return "".join(str(s).split()).lower() if s is not None else ""
def find_cost_col(idx):
    want = {_sq(a) for a in COST_HDRS}
    for h, c in idx.items():
        if _sq(h) in want: return c
    for h, c in idx.items():                  # bien the ten: "... GSC(+ VAT)", "Tong phi GSC +vat"
        s = _sq(h)
        if "trảgsc" in s and "+vat" in s: return c
    return None
def read_bangke(path, sheet, hr):
    """Read a Vietnamese 'bang ke' sheet -> list of claim rows (fee total = 'Phi tra GSC (+vat)')."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    idx = _hdr_index(ws, hr)
    def col(key):
        for a in BK_AL[key]:
            if a in idx: return idx[a]
        return None
    cm = {k: col(k) for k in BK_AL}
    cost_c = find_cost_col(idx)
    if cost_c is None:
        print(f"  ! Sheet '{sheet}': khong tim thay cot 'Phi tra GSC (+vat)' "
              f"-> bo qua canh bao 'Chi phi BT > x% SI'.")
    stt_c, ser_c = _row_key_cols(ws, hr, idx)
    if ser_c is not None:
        print(f"  ! Sheet '{sheet}': cot STT de trong -> nhan dong du lieu bang cot 'IMEI/Serial'.")
    rows = []
    for row in ws.iter_rows(min_row=hr + 1, values_only=False):
        vals = [c.value for c in row]
        if not _is_data_row(vals, stt_c, ser_c): continue
        # read_only: o trong khong ton tai trong XML tra ve EmptyCell (khong co .row),
        # nen lay so dong tu o thuc dau tien chu dung dua vao row[0].
        rnum = next((c.row for c in row if getattr(c, "row", None) is not None), None)
        if rnum is None: continue
        def g(k):
            i = cm[k]; return vals[i] if (i is not None and i < len(vals)) else None
        rows.append(dict(sheet_row=rnum, gsc=stext(g("gsc")), serial=stext(g("serial")),
            phone=stext(g("phone")), mgd=stext(g("mgd")), si=pnum(g("si")), incident=pdate(g("incident")),
            notif=pdate(g("notif")), cause=norm(g("cause")),
            eff_file=pdate(g("eff")), exp_file=pdate(g("exp")), prod=norm(g("prod")),
            cost=pnum(vals[cost_c]) if (cost_c is not None and cost_c < len(vals)) else None))
    wb.close()
    return rows

# Kiem tra dinh dang: o co gia tri nhung pdate/pnum khong doc duoc thi tool bo qua o do (coi nhu trong).
# Ham nay liet ke cac o do de bao cho nguoi dung, khong doi ket qua map.
FMT_DATE = ["incident", "notif", "eff", "exp"]
def _fee_col(h):
    s = _sq(h)
    return "vat" in s or s.startswith("phí")
def format_issues(path, sheet, hr):
    """-> list (sheet_row, header, gia tri goc, loi) cho cot ngay + so tien cua bang ke."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    idx = _hdr_index(ws, hr)
    chk = {}
    for k in FMT_DATE:
        for a in BK_AL[k]:
            if a in idx: chk[idx[a]] = (a, "date"); break
    for a in BK_AL["si"]:
        if a in idx: chk[idx[a]] = (a, "num"); break
    for h, c in idx.items():
        if c not in chk and _fee_col(h): chk[c] = (" ".join(str(h).split()), "num")
    stt_c, ser_c = _row_key_cols(ws, hr, idx)
    out = []
    for row in ws.iter_rows(min_row=hr + 1, values_only=False):
        vals = [c.value for c in row]
        if not _is_data_row(vals, stt_c, ser_c): continue
        rnum = next((c.row for c in row if getattr(c, "row", None) is not None), None)
        for c, (h, kind) in sorted(chk.items()):
            v = vals[c] if c < len(vals) else None
            if v is None or (isinstance(v, str) and v.strip() in ("", "-")): continue
            if kind == "date":
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    out.append((rnum, h, v, "Ngày đang là số, không phải ngày")); continue
                d = pdate(v)
                if d is None: out.append((rnum, h, v, "Không đọc được ngày (cần dd/mm/yyyy)"))
                elif not 2000 <= d.year <= 2100: out.append((rnum, h, v, "Năm không hợp lệ"))
            elif not isinstance(v, (int, float)) and pnum(v) is None:
                out.append((rnum, h, v, "Không phải số"))
    wb.close()
    return out

# ---------- 3. connect Fabric ----------
def connect():
    try:
        import pyodbc
        from azure.identity import (DeviceCodeCredential, TokenCachePersistenceOptions, AuthenticationRecord)
    except ImportError as e:
        die("Missing Python libraries. Launch the tool via 'MAP CLAIM.bat'. (" + str(e) + ")")
    cred = None
    if os.path.exists(AUTH):
        try:
            rec = AuthenticationRecord.deserialize(open(AUTH).read())
            cred = DeviceCodeCredential(client_id=CLIENT,
                cache_persistence_options=TokenCachePersistenceOptions(name="claude_fabric_equip"),
                authentication_record=rec, disable_automatic_authentication=True)
            tok = cred.get_token("https://database.windows.net/.default")
        except Exception:
            cred = None
    if cred is None and HEADLESS:
        # Device-code needs a human to open a browser and type a code. On the server
        # there is none: the flow would block until it times out, holding the request.
        die("The Fabric sign-in on the server has expired, so the tool cannot reach the database.\n"
            "Nam: on the host machine run  py -3.12 map_claim_ri.py --signin  to refresh it "
            "(takes ~1 minute), then retry here.")
    if cred is None:
        print("\n  >>> First run / session expired: you need to SIGN IN to Fabric once.")
        def _prompt(u, c, e):
            if GUI:
                ui_signin(u, c)
            else:
                print(f"\n  >>> Open in a browser: {u}\n  >>> Enter the code: {c}\n      (sign in with a PVI account that has access to the Fabric Equipment workspace)\n", flush=True)
        cred = DeviceCodeCredential(client_id=CLIENT,
            cache_persistence_options=TokenCachePersistenceOptions(name="claude_fabric_equip"),
            prompt_callback=_prompt)
        try:
            rec = cred.authenticate(scopes=["https://database.windows.net/.default"])
            open(AUTH, "w").write(rec.serialize())
            tok = cred.get_token("https://database.windows.net/.default")
        except Exception as e:
            die("Fabric sign-in failed. (" + str(e) + ")")
    t = tok.token.encode("utf-16-le"); st = struct.pack(f"<I{len(t)}s", len(t), t)
    cs = ("DRIVER={ODBC Driver 18 for SQL Server};"
          f"SERVER={SERVER};DATABASE={DBNAME};Encrypt=yes;TrustServerCertificate=no;Connection Timeout=600;")
    try:
        return pyodbc.connect(cs, attrs_before={1256: st}, autocommit=True)
    except Exception as e:
        # Moi loi ODBC deu co chu "[ODBC Driver 18 ...]" nen chi IM002 moi la thieu driver.
        if "IM002" in str(e):
            die("'ODBC Driver 18 for SQL Server' is not installed. See README.txt. (" + str(e) + ")")
        if "18456" in str(e) or "7451" in str(e):
            die("Fabric refused the sign-in (warehouse is updating or the owner session expired). "
                "Nam: open the warehouse at app.fabric.microsoft.com and run one query, then retry. "
                "(" + str(e) + ")")
        die("Cannot connect to Fabric. Check network / account access. (" + str(e) + ")")

# ---------- 4. fetch premium rows + pick the right policy ----------
# We query the BASE table (not vw_Claim_RI_Lookup) on purpose: the view hides the
# cancellation rows (Quantity=-1), and we NEED them to compute net-Quantity so we can
# skip a policy that was cancelled (or cancelled then re-created under a new number).
SRCTBL = "dbo.D_DATA_PREMIUM_EQUIPMENT_V2"
PCOLS = ["IMEI","Phone","RequestID","Policy No","DataSource","Program","Program Coverage",
         "Effective Date","Expired Date","SI","RI Share Pct","Quantity","Premium Gross"]
def _fam_case(colexpr):
    """The product-family CASE from dbo.vw_Claim_Premium_Matched, on one column."""
    c = f"COALESCE({colexpr},'')"
    return ("CASE "
        f"WHEN {c} LIKE '%1-to-1%' OR {c} LIKE '%1D1%' OR {c} LIKE '%MDM%' OR {c} LIKE N'%đổi%' THEN '020413' "
        f"WHEN {c} LIKE '%Crack%' OR {c} LIKE '%NVMH%' OR {c} LIKE N'%Nứt%' THEN '020415' "
        f"WHEN {c} LIKE '%ADLD%' OR {c} LIKE '%AD_%' OR {c} LIKE '% AD%' OR {c} LIKE '%-AD%' OR {c} LIKE N'%Rơi%' "
        f"OR {c} LIKE '%Mandatory%' OR {c} LIKE '%3IN1%' OR {c} LIKE '%IP14%' OR {c} LIKE '%IP1%' THEN '020407' "
        f"WHEN {c} LIKE '%EW%' OR {c} LIKE N'%Bảo hành%' OR {c} LIKE N'%bảo hành%' OR {c} LIKE '%ICT%' "
        f"OR {c} LIKE N'%Tặng%' OR {c} LIKE '%Tang%' THEN '020408' ELSE NULL END")
# same rule as the view: classify by [Program Coverage] first, fall back to [Program].
FAM_SQL = (f"CASE WHEN ({_fam_case('[Program Coverage]')}) IS NOT NULL "
           f"THEN ({_fam_case('[Program Coverage]')}) ELSE ({_fam_case('[Program]')}) END")
FCOLS = PCOLS + ["Family"]
# "Ghi chú" (reinsurance note) column: not loaded to Fabric yet as of 2026-07-06. When it IS
# loaded, an endorsement whose note says "gom sai đơn" is a wrong-consolidation bookkeeping
# reversal (not a real cancel/recreate) and must be IGNORED so the original policy wins.
# The tool auto-detects the column and uses it only if present -> forward-compatible.
NOTE_COL = None
def _detect_note_col(cur):
    global NOTE_COL
    try:
        cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
                    "WHERE TABLE_NAME='D_DATA_PREMIUM_EQUIPMENT_V2' "
                    "AND COLUMN_NAME IN (N'Ghi chú', N'Ghi chu', N'GHI CHÚ', N'Note', N'Notes')")
        r = cur.fetchone(); NOTE_COL = r[0] if r else None
    except Exception:
        NOTE_COL = None
    return NOTE_COL
# ---------- co cau tai: share tung nha tai ----------
# Nguon: dbo.D_RI_STRUCTURE, dung tu hop dong tai trong "Bang ten CT Tai-final _ Gui dot 1.xlsx"
# bang "2. Co cau tai/build_ri_structure_v2.py". Khoa = (DataSource, Program Coverage,
# RI Share Pct) cua chinh dong phi da map -> join 1-1, khong lam no dong.
# LUU Y: cot "RI Share Pct" xuat ra bang TONG share cac nha tai theo HOP DONG, khong phai
# cot [RI Share Pct] cua database (2 cai lech nhau o khoi "chi ghi nha dung dau", vd DB ghi
# 60% = phan Hannover, con Hanoi-re 10% khong co trong data -> hop dong that la 70%).
RI_TBL = "dbo.D_RI_STRUCTURE"
REINSURERS = ["Hannover-re", "Hanoi-re", "TMIV", "Fubon"]
RI_STRUCT = {}      # (ds, pc, share_db) -> {"Hannover-re": 0.6, "Hanoi-re": 0.1}
RI_TOTAL = {}       # (ds, pc, share_db) -> tong share theo hop dong
def _ri_key(ds, pc, share):
    s = pnum(share)
    return (str(ds or "").strip(), str(pc or "").strip(), None if s is None else round(s, 6))
def load_ri_struct(cur):
    """Nap bang co cau tai. Bang chua ton tai -> tra 0, tool van chay (cot nha tai de trong)."""
    RI_STRUCT.clear(); RI_TOTAL.clear()
    try:
        cur.execute(f"SELECT [DataSource],[ProgramCoverage],[RIShareDB],[Reinsurer],[SharePct],"
                    f"[ShareTotal] FROM {RI_TBL}")
        rows = cur.fetchall()
    except Exception:
        return 0
    for ds, pc, shdb, nha, pct, tot in rows:
        k = _ri_key(ds, pc, shdb)
        RI_STRUCT.setdefault(k, {})[str(nha or "").strip()] = pnum(pct)
        RI_TOTAL[k] = pnum(tot)
    return len(RI_TOTAL)

def fetch_in(cur, whereexpr, values):
    out, vals = [], [v for v in values if v]
    cols = FCOLS
    sel = ",".join(f"[{c}]" for c in PCOLS) + f", ({FAM_SQL}) AS [Family]"
    if NOTE_COL:
        sel += f", [{NOTE_COL}] AS [Note]"; cols = FCOLS + ["Note"]
    for i in range(0, len(vals), 200):
        chunk = vals[i:i+200]; ph = ",".join("?"*len(chunk))
        cur.execute(f"SELECT {sel} FROM {SRCTBL} WHERE {whereexpr} IN ({ph})", *chunk)
        out += [dict(zip(cols, r)) for r in cur.fetchall()]
    return out

def ds_group(ds):
    """Business group of a premium row: '1D1' (1 doi 1) vs 'MT' (may tro / khac)."""
    return "1D1" if str(ds or "").strip().upper().startswith("1D1") else "MT"
def _pol_base(pol):
    s = str(pol or ""); i = s.find("/E"); return s[:i] if i > 0 else s
def _pol_enum(pol):
    mm = re.search(r"/E(\d+)", str(pol or "")); return int(mm.group(1)) if mm else 0

def _note_drop(note):
    """Reason a premium row must be IGNORED as a claim map-target, from the RI [Note] column.
       'gomsai' = wrong-consolidation reversal/transfer ; 'test' = test / non-reinsured ("không tái") row."""
    s = str(note or "").strip().lower()
    if not s: return None
    if "gom sai" in s:                                     return "gomsai"
    if "test" in s or "không tái" in s or "khong tai" in s: return "test"
    return None

# KHONG con canh bao "sai san pham" tren cot Luu y HSBT (Nam chot 10/08/2026): don GOI bat buoc
# (Mandatory Program / 3IN1) la MOT don om nhieu quyen loi (roi vo + bao hanh mo rong...), FAM_SQL
# buoc phai xep no vao 1 ho duy nhat (020407 - ADLD, vi tu khoa '%Mandatory%' nam o nhanh ADLD),
# nen moi canh bao lech san pham deu khong dang tin -> bo han 2 note thay vi giai thich tung ca.
# [Family] van duoc dung de CHAM DIEM chon don trong pick_best - chi bo phan ghi chu.
def pick_best(c, cand, group, family):
    """Choose the correct premium row for one claim (mirrors dbo.vw_Claim_Premium_Matched).
       1) keep only rows of the chosen ds-group (1D1/MT);
       2) drop policies whose base net-Quantity <= 0 (cancelled, or cancelled-then-recreated);
       3) among survivors prefer: matching PRODUCT family -> SI = claim SI ->
          covers the loss date -> latest endorsement (max /E) -> latest effective date.
       Returns (best_row|None, tag, fam_ok, void_hit).
       tag: ok / recreate / no-in-group / no-active / no-cand."""
    if not cand:
        return None, "no-cand", True, False, False, False
    pool = [r for r in cand if group is None or ds_group(r["DataSource"]) == group]
    if not pool:
        return None, "no-in-group", True, False, False, False
    # Drop rows flagged in the RI [Note] column that must never be a claim map-target:
    #  - "gom sai đơn": wrong-consolidation reversal/transfer -> ignore so the real original policy wins.
    #  - test / "không tái": dummy or non-reinsured rows -> not a valid policy to map a claim onto.
    void_hit = any(_note_drop(r.get("Note")) == "gomsai" for r in pool)
    test_hit = any(_note_drop(r.get("Note")) == "test"   for r in pool)
    if void_hit or test_hit:
        pool = [r for r in pool if _note_drop(r.get("Note")) is None]
        if not pool:
            return None, "no-active", True, void_hit, test_hit, False
    net = {}
    for r in pool:
        b = _pol_base(r["Policy No"]); net[b] = net.get(b, 0.0) + (pnum(r["Quantity"]) or 0.0)
    active = {b for b, q in net.items() if q > 0}
    issue = [r for r in pool if (pnum(r["Premium Gross"]) or 0) > 0 and (pnum(r["Quantity"]) or 0) != -1]
    live  = [r for r in issue if _pol_base(r["Policy No"]) in active]
    tag = "ok"
    if live:
        pool2 = live
        if any(q <= 0 for q in net.values()): tag = "recreate"   # some base was cancelled
    elif issue:
        pool2, tag = issue, "no-active"                          # everything cancelled -> flag
    else:
        return None, "no-active", True, void_hit, test_hit, False
    # File's own policy dates ("Ngày hiệu lực"/"Ngày hết hiệu lực") disambiguate when one IMEI
    # has policies in >1 programme (e.g. bán lẻ vs tặng): match them to the DB Effective/Expired.
    cef, cxp = c.get("eff_file"), c.get("exp_file")
    def _dscore(r):
        """2 = both file dates equal this row's Eff+Exp; 1 = one of them; 0 = no file date matches."""
        reff, rexp = pdate(r["Effective Date"]), pdate(r["Expired Date"])
        if cef and reff and cef == reff:
            return 2 if (cxp and rexp and cxp == rexp) else 1
        if cxp and rexp and cxp == rexp:
            return 1
        return 0
    def key(r):
        fam = 1 if (family is None or r.get("Family") == family) else 0
        rsi = pnum(r["SI"]); sim = 1 if (c["si"] is not None and rsi is not None and abs(rsi - c["si"]) < 1) else 0
        reff, rexp = pdate(r["Effective Date"]), pdate(r["Expired Date"])
        cov = 1 if (reff and rexp and c["incident"] and reff <= c["incident"] <= rexp) else 0
        # dm ranks FIRST but is 0 for everyone when the file gives no matching date -> then the
        # tuple order is unchanged and behaviour is identical to before this feature.
        dm = _dscore(r)
        return (dm, fam, sim, cov, _pol_enum(r["Policy No"]), reff.toordinal() if reff else 0)
    best = max(pool2, key=key)
    fam_ok = (family is None or best.get("Family") == family)
    # flag only when the file dates actually decided across >1 distinct programme (the user's case)
    progs = {(norm(r.get("Program Coverage")) or norm(r.get("Program")) or "") for r in pool2}
    date_hit = _dscore(best) > 0 and len(progs) >= 2
    return best, tag, fam_ok, void_hit, test_hit, date_hit

# Canh bao "Khac chuong trinh" (Nam yeu cau lai 29/09/2026, sau ca EW map ra ADLD PLUS).
# KHONG dung [Family] (ban 10/08 bi bo vi goi 3IN1/Mandatory bi xep vao ADLD -> bao nham).
# So loai quyen loi theo TIEN TO [Program Coverage] (ADLD_/EW_/CS_, sach tren toan bang phi)
# voi "Ten san pham" cua tung dong trong bang ke; khong co cot do thi so voi san pham cua dot.
# Don goi (3IN1 / Mandatory) om nhieu quyen loi -> khong bao gio canh bao.
_COV_LBL = {"020407": "ADLD", "020408": "EW", "020415": "Nứt màn hình"}
_PKG = ("3IN1", "3 IN 1", "MANDATORY")
def _cov_type(best):
    pc = (norm(best.get("Program Coverage")) or "").upper()
    for pre, code in (("ADLD", "020407"), ("EW", "020408"), ("CS", "020415")):
        if pc.startswith(pre): return code
    return None
def _file_type(prod):
    s = (prod or "").upper()
    if not s or any(k in s for k in _PKG) or "1D1" in s or "1 ĐỔI 1" in s: return None
    if "NVMH" in s or "CRACK" in s or "NỨT" in s or "MÀN HÌNH" in s: return "020415"
    if "ADLD" in s or "RƠI" in s or re.search(r"\bAD\b", s): return "020407"
    if re.search(r"\bEW\b", s) or "BẢO HÀNH" in s: return "020408"
    return None
def prog_mismatch(c, best, family):
    if not best: return ""
    prog = (norm(best.get("Program")) or "").upper()
    if any(k in prog for k in _PKG): return ""
    got = _cov_type(best)
    want = _file_type(c.get("prod")) if c.get("prod") else None
    src = "hồ sơ ghi '%s'" % c["prod"] if want else None
    if want is None and not c.get("prod") and family in _COV_LBL:
        want, src = family, "đợt %s" % _COV_LBL[family]
    if not want or not got or want == got: return ""
    return (f"Khác chương trình: {src} nhưng đơn map là {_COV_LBL[got]} "
            f"({norm(best.get('Program')) or ''}), kiểm tra lại hồ sơ")

def match(cur, claims, group=None, family=None):
    _detect_note_col(cur)        # use the reinsurance "Ghi chú" column if it exists on the table
    load_ri_struct(cur)          # co cau tai (share tung nha tai) tu dbo.D_RI_STRUCTURE
    # priority: 1) IMEI/Serial  2) Mã giao dịch = RequestID  3) Phone (last 9)
    serials = sorted({c["serial"] for c in claims if c["serial"]})
    by_ser = {}
    for d in fetch_in(cur, "LTRIM(RTRIM([IMEI]))", serials):
        by_ser.setdefault(str(d["IMEI"]).strip(), []).append(d)
    def in_ser(c): return bool(c["serial"] and c["serial"] in by_ser)
    need_req = sorted({c.get("mgd") for c in claims if c.get("mgd") and not in_ser(c)})
    by_req = {}
    for d in fetch_in(cur, "LTRIM(RTRIM([RequestID]))", need_req):
        k = str(d["RequestID"]).strip() if d["RequestID"] else None
        if k: by_req.setdefault(k, []).append(d)
    def in_req(c): return bool(c.get("mgd") and c["mgd"] in by_req)
    need = sorted({c["phone9"] for c in claims if c["phone9"] and not in_ser(c) and not in_req(c)})
    by_ph = {}
    for d in fetch_in(cur, "RIGHT(LTRIM(RTRIM([Phone])),9)", need):
        k = str(d["Phone"]).strip()[-9:] if d["Phone"] else None
        if k: by_ph.setdefault(k, []).append(d)
    out = []
    for c in claims:
        cand, key = [], None
        if in_ser(c):                               cand, key = by_ser[c["serial"]], "Serial"
        elif in_req(c):                             cand, key = by_req[c["mgd"]], "MaGiaoDich"
        elif c["phone9"] and c["phone9"] in by_ph:  cand, key = by_ph[c["phone9"]], "Phone"
        # fam_ok chi con de tham khao: khong con sinh note "sai san pham" nua (10/08/2026)
        best, tag, _fam_ok, void_hit, test_hit, date_hit = pick_best(c, cand, group, family)
        notes = []
        if best is None:
            if void_hit or test_hit:
                notes.append("Thiết bị chỉ có đơn gom sai/test/không tái, không map, cần kiểm tra")
            elif tag == "no-in-group":
                notes.append(f"Chưa map được policy trong nhóm {group} (thiết bị có đơn ở nhóm khác, kiểm tra)")
        else:
            pm = prog_mismatch(c, best, family)
            if pm: notes.append(pm)
            if date_hit:
                notes.append("Đã chọn đơn theo Ngày hiệu lực/hết hiệu lực khớp file (thiết bị có nhiều chương trình)")
            if void_hit:
                notes.append("Đã bỏ endor 'gom sai đơn' theo Ghi chú, map về đơn gốc")
            if test_hit:
                notes.append("Đã bỏ đơn test/không tái theo Ghi chú, map về đơn còn hiệu lực")
            if not void_hit and not test_hit:
                if tag == "recreate":
                    notes.append(f"Đơn gốc đã hủy, đã map vào đơn còn hiệu lực ({best['Policy No']})")
                elif tag == "no-active":
                    notes.append("Mọi đơn của thiết bị đều đã hủy, tạm map, cần kiểm tra")
        out.append(dict(c=c, key=(key if best else None), best=best, note_extra=" | ".join(notes),
                        distpol=len({str(r["Policy No"]).strip() for r in cand})))
    return out

# ---------- 5. Luu y HSBT ----------
def hsbt_note(c, best):
    n = []
    eff = pdate(best["Effective Date"]) if best else None
    exp = pdate(best["Expired Date"])   if best else None
    si  = pnum(best["SI"]) if best else c["si"]
    loss, notif = c["incident"], c["notif"]
    if not best:
        n.append("Chưa map được policy")
    if loss and eff and exp:
        if loss < eff or loss > exp:
            n.append(f"Ngày TT NGOÀI thời hạn BH (TT {loss:%d/%m/%Y}, HLBH {eff:%d/%m/%Y}–{exp:%d/%m/%Y})")
        else:
            ds, de = (loss - eff).days, (exp - loss).days
            if ds <= 3: n.append(f"Ngày TT sát ngày bắt đầu HLBH ({ds} ngày)")
            if de <= 3: n.append(f"Ngày TT sát ngày hết hạn HLBH ({de} ngày)")
    if notif and eff and exp and (notif < eff or notif > exp):
        n.append(f"Ngày báo TT ngoài thời hạn BH (báo {notif:%d/%m/%Y})")
    if notif and loss:
        dd = (notif - loss).days
        if dd > 30: n.append(f"Báo TT trễ {dd} ngày so với ngày TT (>30)")
    cause = ((c["cause"] or "") + " " + (c["detail"] or "")).lower()
    if "pin" in cause: n.append("Nguyên nhân tổn thất có 'Pin'")
    if c["cost"] and si and si > 0:
        r = c["cost"] / si
        if   r > 1.0:  n.append(f"Chi phí BT > 100% SI ({r*100:.0f}%)")
        elif r > 0.9:  n.append(f"Chi phí BT > 90% SI ({r*100:.0f}%)")
        elif r > 0.75: n.append(f"Chi phí BT > 75% SI ({r*100:.0f}%)")
    return " | ".join(n)

# ---------- 6. write output ----------
NAVY="183665"; BLUE="2E5496"; GREY="595959"; ORANGE="C55A11"
def _row_values(m):
    """7 truong map + 4 nha tai + key + sicheck + hsbt cho 1 dong claim."""
    best = m["best"]
    _extra = m.get("note_extra", "")
    _base_note = hsbt_note(m["c"], best)
    _hsbt = _extra if (not best and _extra) else " | ".join(x for x in (_extra, _base_note) if x)
    v = {"Policy No":"", "Program":"", "Program Coverage":"",
         "Effective Date":None, "Expired Date":None, "SI":None, "RI Share Pct":None,
         "key": (m["key"] if best else "NO MATCH"), "hsbt": _hsbt, "sicheck": ""}
    for _nha in REINSURERS: v[_nha] = None
    if best:
        cs = m["c"]["si"]; ps = pnum(best["SI"])
        v["sicheck"] = "" if (cs is None or ps is None) else ("OK" if abs(ps-cs) < 1 else "DIFF")
        v.update({"Policy No":best["Policy No"], "Program":norm(best["Program"]) or "",
            "Program Coverage":best["Program Coverage"],
            "Effective Date":pdate(best["Effective Date"]), "Expired Date":pdate(best["Expired Date"]),
            "SI":pnum(best["SI"]), "RI Share Pct":pnum(best["RI Share Pct"])})
        # Co cau tai: share tung nha + [RI Share Pct] = TONG share theo hop dong tai.
        k = _ri_key(best.get("DataSource"), best.get("Program Coverage"), best.get("RI Share Pct"))
        sp = RI_STRUCT.get(k)
        if sp:
            for _nha in REINSURERS: v[_nha] = sp.get(_nha)
            if RI_TOTAL.get(k) is not None: v["RI Share Pct"] = RI_TOTAL[k]
        elif RI_TOTAL:   # bang co cau da nap nhung to hop nay chua co trong hop dong
            v["hsbt"] = " | ".join(x for x in (v["hsbt"],
                "Chưa có cơ cấu tái trong hợp đồng - tỷ lệ lấy theo database") if x)
    return v

def fmt_d(d):
    return d.strftime("%d/%m/%Y") if isinstance(d, datetime.date) else ""
def fmt_pct(p):
    return "" if p is None else f"{round(p*100,1):g}%"
# "RI Share Pct" = tong share cac nha tai theo hop dong; 4 cot sau la phan tung nha.
OUT_HEADERS = ["Policy No","Program","Program Coverage","Effective Date","Expired Date",
               "SI","RI Share Pct"] + REINSURERS + ["Lưu ý HSBT"]
OUT_WIDTHS  = [22,28,24,13,13,14,11,   12,10,8,8,   52]
EMPTY_V = {"Policy No":"","Program":"","Program Coverage":"","Effective Date":None,"Expired Date":None,
           "SI":None,"RI Share Pct":None,"key":"NO MATCH","sicheck":"","hsbt":"Chưa map được policy",
           **{n: None for n in REINSURERS}}
def _int_or_none(v):
    if v is None: return None
    f = pnum(v)
    return None if f is None else (int(f) if float(f).is_integer() else f)
def out_cells(v):
    return [(v.get("Policy No",""),"s"), (v.get("Program",""),"s"), (v.get("Program Coverage",""),"s"),
            (xldate(v.get("Effective Date")),"d"), (xldate(v.get("Expired Date")),"d"),
            (_int_or_none(v.get("SI")),"money"), (fmt_pct(v.get("RI Share Pct")),"s")] + \
           [(fmt_pct(v.get(n)),"s") for n in REINSURERS] + \
           [(v.get("hsbt",""),"s")]

LINE_PT = 18.0   # chieu cao 1 dong chu (dong bang ke goc cao 36pt ~ 2 dong)
def _row_min_height(v):
    """Uoc luong chieu cao toi thieu (pt) de chu trong cac cot moi KHONG bi cat.
       Tra ve None neu <=2 dong (giu nguyen chieu cao goc)."""
    mx = 1
    for (val, _k), w in zip(out_cells(v), OUT_WIDTHS):
        s = "" if val is None else str(val)
        if not s: continue
        segs = s.split(" | ") if " | " in s else [s]   # moi flag "Luu y" thuong xuong dong rieng
        lines = sum(max(1, -(-len(seg) // int(w))) for seg in segs)
        if lines > mx: mx = lines
    return mx * LINE_PT if mx > 2 else None

# ---- surgical: chen cot vao XML cua sheet, GIU NGUYEN byte moi phan khac (logo/pivot/table) ----
def _xesc(s):
    return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
def _cell_xml(ref, s, value, kind):
    if value is None or value == "":
        return f'<c r="{ref}" s="{s}"/>'
    if kind in ("n", "money", "d"):                # numeric payload (money = comma fmt, d = serial date)
        return f'<c r="{ref}" s="{s}"><v>{value}</v></c>'
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{_xesc(value)}</t></is></c>'
def _get_s(xml, ref):
    m = re.search(r'<c r="%s"[^>]*\bs="(\d+)"' % ref, xml)
    return m.group(1) if m else "0"
def _text_num_styles(xml, rownum):
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rownum, xml, re.S)
    body = m.group(1) if m else ""
    s_txt = s_num = None
    for cm in re.finditer(r'<c\b([^>]*?)/?>', body):
        a = cm.group(1)
        sm = re.search(r'\bs="(\d+)"', a); s = sm.group(1) if sm else "0"
        tm = re.search(r'\bt="(\w+)"', a)
        if tm and tm.group(1) in ("s","str","inlineStr"):
            if s_txt is None: s_txt = s
        elif not tm and "<v>" in body:
            if s_num is None: s_num = s
    return (s_txt or "0", s_num or s_txt or "0")

def _clone_xf_numfmt(xf, numfmt_id):
    """Clone a cellXf, forcing its number format (keeps font/fill/border/alignment)."""
    if re.search(r'\bnumFmtId="\d+"', xf):
        xf = re.sub(r'\bnumFmtId="\d+"', 'numFmtId="%d"' % numfmt_id, xf, count=1)
    else:
        xf = re.sub(r'<xf\b', '<xf numFmtId="%d"' % numfmt_id, xf, count=1)
    if re.search(r'\bapplyNumberFormat="\d+"', xf):
        xf = re.sub(r'\bapplyNumberFormat="\d+"', 'applyNumberFormat="1"', xf, count=1)
    else:
        xf = re.sub(r'<xf\b', '<xf applyNumberFormat="1"', xf, count=1)
    return xf
def _last_content_col(xml, rownum):
    """1-based index of the last cell in <row rownum> that actually holds a value (ignores empty cells)."""
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % rownum, xml, re.S)
    if not m: return 0
    maxc = 0
    for cm in re.finditer(r'<c r="([A-Z]+)\d+"[^>]*?(/>|>(.*?)</c>)', m.group(1), re.S):
        ref, closing, inner = cm.group(1), cm.group(2), cm.group(3)
        if closing == "/>": continue
        if inner and ("<v>" in inner or "<t" in inner or "<is>" in inner):
            ci = openpyxl.utils.column_index_from_string(ref)
            if ci > maxc: maxc = ci
    return maxc
def _ensure_date_money_styles(styles_xml, base_idx):
    """Append 2 cellXfs (date dd/mm/yyyy + comma number) cloned from cellXfs[base_idx].
       Return (new_styles_xml, s_date_index, s_money_index) as string indices."""
    xfs = _xf_list(_inner(styles_xml, "cellXfs"))
    base = xfs[base_idx] if 0 <= base_idx < len(xfs) else '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    ids = [int(x) for x in re.findall(r'<numFmt numFmtId="(\d+)"', styles_xml)]
    date_id = max([163] + ids) + 1                         # custom numFmt id (>=164); 3 = builtin "#,##0"
    n = len(xfs)
    styles_xml = _add_container(styles_xml, "cellXfs",
                                _clone_xf_numfmt(base, date_id) + _clone_xf_numfmt(base, 3), 2)
    styles_xml = re.sub(r'<numFmts\b[^>]*/>', '', styles_xml)
    styles_xml = _add_container(styles_xml, "numFmts",
                                '<numFmt numFmtId="%d" formatCode="dd/mm/yyyy"/>' % date_id, 1, before_tag="fonts")
    return styles_xml, str(n), str(n + 1)

def surgical_write(in_path, edits):
    """edits: list cua (sheet_name, header_row, {rownum: values_dict}). Chen cot moi vao sau cot cuoi."""
    out_path = os.path.splitext(in_path)[0] + "_MAPPED.xlsx"
    try:
        zin = zipfile.ZipFile(in_path)
    except Exception as e:
        die(f"Cannot open the input file for writing. ({e})")
    names = zin.namelist(); parts = {n: zin.read(n) for n in names}; zin.close()
    wbxml = parts["xl/workbook.xml"].decode("utf-8")
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    ncol = len(OUT_HEADERS)
    s_date = s_money = None
    for sheet_name, hdr_row, rowvals in edits:
        if not rowvals: continue
        sm = re.search(r'<sheet[^>]*name="%s"[^>]*r:id="(rId\d+)"' % re.escape(sheet_name), wbxml)
        if not sm: continue
        tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % sm.group(1), rels).group(1).replace("\\", "/")
        sp = tgt[1:] if tgt.startswith("/") else "xl/" + tgt
        xml = parts[sp].decode("utf-8")
        dm = re.search(r'<dimension ref="A1:([A-Z]+)(\d+)"\s*/>', xml)
        # append AFTER the last column that actually has a header (not trailing empty cols counted in the
        # dimension) - a blank column left mid-range would break the AutoFilter dropdowns
        maxcol = _last_content_col(xml, hdr_row) or (openpyxl.utils.column_index_from_string(dm.group(1)) if dm else 1)
        first = maxcol + 1; last_n = first + ncol - 1
        last_letter = get_column_letter(last_n)
        s_hdr = _get_s(xml, "%s%d" % (get_column_letter(1), hdr_row))
        s_txt, s_num = _text_num_styles(xml, min(rowvals.keys()))
        if s_date is None:                       # create date + comma-number styles once (shared styles.xml)
            stx, s_date, s_money = _ensure_date_money_styles(parts["xl/styles.xml"].decode("utf-8"), int(s_num))
            parts["xl/styles.xml"] = stx.encode("utf-8")
        kstyle = {"s": s_txt, "n": s_num, "d": s_date, "money": s_money}
        inject = {hdr_row: "".join(_cell_xml("%s%d" % (get_column_letter(first+i), hdr_row), s_hdr, OUT_HEADERS[i], "s")
                                   for i in range(ncol))}
        row_h = {}
        for rn, v in rowvals.items():
            cs = out_cells(v)
            inject[rn] = "".join(_cell_xml("%s%d" % (get_column_letter(first+i), rn),
                                           kstyle.get(cs[i][1], s_txt), cs[i][0], cs[i][1])
                                 for i in range(ncol))
            h = _row_min_height(v)
            if h: row_h[rn] = h
        def _strip_ge_first(body, rn):
            # remove any pre-existing (usually empty, styled) cells at/after the insert column,
            # else appending mapped cells there creates DUPLICATE cell refs -> Excel repairs & drops the filter
            return re.sub(r'<c r="([A-Z]+)%d"[^>]*?(?:/>|>.*?</c>)' % rn,
                          lambda m: "" if openpyxl.utils.column_index_from_string(m.group(1)) >= first else m.group(0),
                          body, flags=re.S)
        def repl(mt):
            rn = int(mt.group("rn"))
            if rn not in inject: return mt.group(0)
            body = _strip_ge_first(mt.group("body"), rn)
            opentag = re.sub(r'spans="\d+:\d+"', 'spans="1:%d"' % last_n, mt.group("open"))
            if rn in row_h:                      # noi cao dong cho du chu (chi noi, khong thu nho)
                cur = re.search(r'\bht="([\d.]+)"', opentag)
                newh = max(float(cur.group(1)) if cur else 0.0, row_h[rn])
                if not cur or newh > float(cur.group(1)):
                    opentag = re.sub(r'\s+ht="[^"]*"', '', opentag)
                    opentag = re.sub(r'\s+customHeight="[^"]*"', '', opentag)
                    opentag = opentag[:-1] + ' ht="%g" customHeight="1">' % newh
            return opentag + body + inject[rn] + "</row>"
        xml = re.sub(r'(?P<open><row r="(?P<rn>\d+)"[^>]*>)(?P<body>.*?)</row>', repl, xml, flags=re.S)
        if dm:
            xml = re.sub(r'<dimension ref="A1:[A-Z]+\d+"\s*/>',
                         '<dimension ref="A1:%s%s"/>' % (last_letter, dm.group(2)), xml, count=1)
        colstr = "".join('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (first+i, first+i, OUT_WIDTHS[i])
                         for i in range(ncol))
        if "<cols>" in xml:
            xml = xml.replace("</cols>", colstr + "</cols>", 1)
        else:
            xml = re.sub(r'(<sheetData)', "<cols>" + colstr + "</cols>" + r"\1", xml, count=1)
        # AutoFilter: extend to cover the mapped columns. Excel keeps the filter range in BOTH the sheet
        # <autoFilter> AND the workbook _FilterDatabase name; if they disagree the dropdowns stop working.
        old_af = re.search(r'<autoFilter\b[^>]*?(?:/>|>.*?</autoFilter>)', xml, re.S)
        uidm = re.search(r'\bxr:uid="[^"]*"', old_af.group(0)) if old_af else None
        dn = re.search(r"<definedName name=\"_xlnm\._FilterDatabase\"[^>]*>'?%s'?!"
                       r"\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)</definedName>" % re.escape(sheet_name), wbxml)
        if dn:
            c0, r0, c1old, r1 = dn.group(1), dn.group(2), dn.group(3), dn.group(4)
        else:
            c0, r0, c1old, r1 = "A", str(hdr_row), None, str(max(rowvals.keys()))
        af = '<autoFilter ref="%s%s:%s%s"%s/>' % (c0, r0, last_letter, r1,
                                                  (" " + uidm.group(0)) if uidm else "")
        if old_af:
            xml = xml.replace(old_af.group(0), af, 1)
        else:
            xml = xml.replace("</sheetData>", "</sheetData>" + af, 1)
        if dn:                                   # keep workbook filter range in sync (extend end column)
            new_dn = dn.group(0).replace("$%s$%s:$%s$%s" % (c0, r0, c1old, r1),
                                         "$%s$%s:$%s$%s" % (c0, r0, last_letter, r1))
            wbxml = wbxml.replace(dn.group(0), new_dn, 1)
        parts[sp] = xml.encode("utf-8")
    parts["xl/workbook.xml"] = wbxml.encode("utf-8")
    try:
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
            for n in names:
                zo.writestr(n, parts[n])
    except PermissionError:
        die(f"Cannot write '{os.path.basename(out_path)}' (is it open in Excel?). Close it and run again.")
    return out_path

# ---------- 7. build PIAS input file from the mapped statement sheet ----------
PIAS_CANBO = "45.CB030007"
PIAS_TPL   = os.path.join(HERE, "Template nhap HSBT Pias.xlsx")   # the desktop app points this at app\data
def gen_pias(mapped_path, bk_sheet, bk_hdr):
    """Doc sheet bang ke da map, dung form PIAS tu template -> tra ve (bytes_xlsx, n_rows).
       (Khong luu file rieng: noi dung nay duoc ghep thanh sheet 'Pias' trong file MAPPED.)"""
    tpl = PIAS_TPL
    if not os.path.exists(tpl):
        print("  (skip PIAS: 'Template nhap HSBT Pias.xlsx' not found in folder)"); return None
    from copy import copy as _copy
    def nz(s): return re.sub(r"\s+", " ", str(s)).strip() if s is not None else ""
    wb = openpyxl.load_workbook(mapped_path, data_only=True)
    ws = wb[bk_sheet]; HR = bk_hdr
    ordered = [(nz(ws.cell(row=HR, column=c).value), c) for c in range(1, ws.max_column+1)
               if ws.cell(row=HR, column=c).value is not None]
    first = {}
    for nm, c in ordered:
        if nm and nm not in first: first[nm] = c
    def ex(*names):
        for n in names:
            if n in first: return first[n]
        return None
    def fc(*subs):
        for nm, c in ordered:
            nl = nm.lower()
            if all(s.lower() in nl for s in subs): return c
        return None
    def vat_after(prec):                 # VAT column right after a fee column (naming varies by batch)
        if not prec: return None
        for nm, c in ordered:
            if c > prec and "vat" in nm.lower(): return c
        return None
    lab_pre  = fc("nhân công", "(-vat)")
    part_pre = fc("linh kiện", "sau mt", "(-vat)") or fc("linh kiện", "(-vat)")
    C = dict(stt=ex("STT"), policy=ex("Policy No"), serial=ex("IMEI/Serial"), model=ex("Model"),
             mgd=ex("Mã giao dịch", "Ma giao dich"),
             dol=ex("Ngày tổn thất"), notif=ex("Ngày thông báo"), invno=ex("Số HĐ"), invdate=ex("Ngày HĐ"),
             lab_pre=lab_pre,  lab_vat=ex("VAT 8%") or vat_after(lab_pre),
             part_pre=part_pre, part_vat=ex("VAT 10% & 8%") or vat_after(part_pre))
    if not C["policy"] or not (C["stt"] or C["serial"]):
        print("  (skip PIAS: cannot find STT / Policy No columns)"); return None
    # cung quy tac nhan dong voi read_bangke: cot STT rong/thieu TOAN sheet thi nhan dong theo IMEI/Serial
    def _num(v): return isinstance(v, (int, float)) and v
    use_stt = bool(C["stt"]) and any(_num(ws.cell(row=r, column=C["stt"]).value)
                                     for r in range(HR+1, ws.max_row+1))
    if not use_stt: print("  (PIAS: STT column empty, rows taken by IMEI/Serial)")
    def gv(r, k): return ws.cell(row=r, column=C[k]).value if C[k] else None
    # 1) one record per claim row (keep the fee components so we can sum by contract)
    from collections import OrderedDict
    claims = []
    for r in range(HR+1, ws.max_row+1):
        if use_stt:
            if not _num(ws.cell(row=r, column=C["stt"]).value): continue
        elif not str(gv(r, "serial") or "").strip(): continue
        claims.append(dict(                          # stext: IMEI dang so -> text, khong de Excel hien E+14
            policy=nz(stext(gv(r,"policy"))), serial=stext(gv(r,"serial")), mgd=nz(gv(r,"mgd")),
            dol=pdate(gv(r,"dol")), notif=pdate(gv(r,"notif")),
            lab_pre=pnum(gv(r,"lab_pre")) or 0, lab_vat=pnum(gv(r,"lab_vat")) or 0,
            part_pre=pnum(gv(r,"part_pre")) or 0, part_vat=pnum(gv(r,"part_vat")) or 0))
    # 2) group by contract (Policy No) so it is NOT duplicated; sum the fees. Rows with no
    #    policy never merge (each keeps its own group). Representative IMEI = the one with a
    #    BLANK "Mã giao dịch" and the LATEST date of loss (fallback: latest dol overall).
    groups = OrderedDict()
    for i, c in enumerate(claims):
        key = c["policy"] if c["policy"] else ("__nopol__", i)
        groups.setdefault(key, []).append(c)
    def _rep_key(m):
        return (1 if not m["mgd"] else 0, m["dol"] or datetime.date.min)
    # 3) inside a contract, split each fee type by VAT rate (8% vs 10%): summing mixed rates
    #    gives a blended 8.3% that rounds to 0.08 and hides the 10% claim. One PIAS row per
    #    (policy, fee type, rate). IMEI = contract representative (as before); a split-off
    #    rate that does not contain it uses its own best claim.
    rows = []
    for members in groups.values():
        prep = max(members, key=_rep_key)
        for pk, vk, dien in [("lab_pre","lab_vat","phí nhân công"),
                             ("part_pre","part_vat","phí sửa chữa/ thay thế")]:
            by_rate = {}
            for m in members:
                if m[pk] <= 0: continue
                by_rate.setdefault(round(m[vk]/m[pk], 2), []).append(m)
            for rate in sorted(by_rate):
                sub = by_rate[rate]
                rep = prep if len(by_rate) == 1 or any(m is prep for m in sub) else max(sub, key=_rep_key)
                pre = sum(m[pk] for m in sub); vat = sum(m[vk] for m in sub)
                rows.append(dict(policy=rep["policy"], serial=rep["serial"], model="",
                                 dol=rep["dol"], notif=rep["notif"],
                                 pre=pre, vat=vat, rate=round(vat/pre, 2), dien=dien))
    # ---- fill template (tidy: aligned rows & columns, comma numbers, real dates) ----
    tb = openpyxl.load_workbook(tpl); ps = tb[tb.sheetnames[0]]
    sty = {c: _copy(ps.cell(row=4, column=c)._style) for c in range(1, 21)}
    for r in range(4, ps.max_row+1):
        for c in range(1, 21): ps.cell(row=r, column=c).value = None
    def sc(r, c, v):
        cell = ps.cell(row=r, column=c, value=v); cell._style = _copy(sty[c]); return cell
    MONEY = "#,##0"; DATEF = "dd/mm/yyyy"
    BOX = Border(left=Side(style="thin"), right=Side(style="thin"),
                 top=Side(style="thin"), bottom=Side(style="thin"))   # full box -> every column separated
    for r in (1, 2, 3):                              # header rows: box every cell so all column headers separate
        for c in range(1, 21): ps.cell(row=r, column=c).border = BOX
    def al(cell, horiz):                              # data cells: single line (no wrap), vertically centered
        cell.alignment = Alignment(horizontal=horiz, vertical="center", wrap_text=False); return cell
    def num(cell, fmt, horiz):
        cell.number_format = fmt; return al(cell, horiz)
    rr = 4
    for i, p in enumerate(rows, 1):
        for c in range(1, 21):                       # base style (fill/font/align) for blanks too
            ps.cell(row=rr, column=c)._style = _copy(sty[c])
        al(sc(rr, 1, i), "center")
        al(sc(rr, 3, p["policy"]), "left")
        num(sc(rr, 4, p["serial"]), "@", "left")
        num(sc(rr, 5, p["pre"]),  MONEY, "right")
        num(sc(rr, 6, p["rate"]), "0.00", "center")
        num(sc(rr, 7, p["vat"]),  MONEY, "right")
        num(sc(rr, 9, p["dol"]),  DATEF, "center")
        al(sc(rr, 14, PIAS_CANBO), "center")
        al(sc(rr, 20, p["dien"]), "left")
        # col8 Doi tuong TT, col10 Ngay thong bao, col17 So HD, col18 Ngay HD -> left blank for PIAS entry
        for c in range(1, 21): ps.cell(row=rr, column=c).border = BOX   # LAST: full box on every cell
        ps.row_dimensions[rr].height = 30            # neat, uniform row height (was 94.5pt sample)
        rr += 1
    for c in (5, 7):                                 # SUBTOTAL row: comma format, bold, right-aligned
        cl = get_column_letter(c)
        cell = ps.cell(row=rr, column=c, value=f"=SUBTOTAL(9,{cl}4:{cl}{rr-1})")
        cell._style = _copy(sty[c]); cell.font = Font(bold=True)
        cell.number_format = MONEY; cell.alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 21):                           # close the table: border the whole SUBTOTAL row
        cell = ps.cell(row=rr, column=c)
        if c not in (5, 7): cell._style = _copy(sty[c])
        cell.border = BOX
    ps.column_dimensions["A"].width = 6              # STT
    ps.column_dimensions["D"].width = 26             # IMEI/Serial - wide enough so long serials don't wrap
    ps.column_dimensions["N"].width = 16             # Cán bộ (45.CB030007)
    if not ps.column_dimensions["O"].width: ps.column_dimensions["O"].width = 12
    bio = io.BytesIO(); tb.save(bio)
    return bio.getvalue(), len(rows)

# ---------- 7b. ghep sheet 'Pias' vao file MAPPED (graft o muc XML, giu nguyen pivot/logo/table) ----------
PIAS_SHEET_NAME = "Pias"
def _xf_list(inner):
    # tach tung <xf .../> hoac <xf ...>...</xf> (the mo cua xf khong chua ky tu '>')
    return re.findall(r'<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>', inner or "", re.S)
def _inner(xml, tag):
    m = re.search(r'<%s\b[^>]*>(.*?)</%s>' % (tag, tag), xml, re.S)
    return m.group(1) if m else None
def _count(xml, tag):
    m = re.search(r'<%s\b[^>]*?\bcount="(\d+)"' % tag, xml)
    return int(m.group(1)) if m else 0
def _remap_xf(xf, nm, foff, filloff, boff, soff=None):
    xf = re.sub(r'numFmtId="(\d+)"', lambda m: 'numFmtId="%d"' % nm.get(int(m.group(1)), int(m.group(1))), xf)
    xf = re.sub(r'fontId="(\d+)"',   lambda m: 'fontId="%d"'   % (int(m.group(1)) + foff), xf)
    xf = re.sub(r'fillId="(\d+)"',   lambda m: 'fillId="%d"'   % (int(m.group(1)) + filloff), xf)
    xf = re.sub(r'borderId="(\d+)"', lambda m: 'borderId="%d"' % (int(m.group(1)) + boff), xf)
    if soff is not None:
        xf = re.sub(r'xfId="(\d+)"', lambda m: 'xfId="%d"' % (int(m.group(1)) + soff), xf)
    return xf
def _add_container(xml, tag, addition, addcount, before_tag=None):
    """Chen 'addition' vao cuoi container <tag> va cong don count len addcount.
       Neu container chua co (va co before_tag) -> tao moi truoc <before_tag>."""
    if ('</%s>' % tag) in xml:
        xml = re.sub(r'<%s\b[^>]*>' % tag,
                     lambda m: re.sub(r'count="(\d+)"',
                                      lambda c: 'count="%d"' % (int(c.group(1)) + addcount),
                                      m.group(0), count=1),
                     xml, count=1)
        xml = xml.replace('</%s>' % tag, addition + '</%s>' % tag, 1)
    elif before_tag:
        block = '<%s count="%d">%s</%s>' % (tag, addcount, addition, tag)
        xml = re.sub(r'<%s\b' % before_tag, block + '<%s' % before_tag, xml, count=1)
    return xml
def merge_styles(mp, ps):
    """Gop styles.xml cua PIAS (ps) vao cua MAPPED (mp). Tra ve (mp_moi, xoff) — xoff la so
       cellXfs cua MAPPED de cong vao moi chi so style 's' tren sheet PIAS."""
    foff    = _count(mp, "fonts")
    filloff = _count(mp, "fills")
    boff    = _count(mp, "borders")
    soff    = _count(mp, "cellStyleXfs")
    xoff    = _count(mp, "cellXfs")
    mp = re.sub(r'<numFmts\b[^>]*/>', '', mp)               # bo <numFmts/> rong (neu co) de them moi
    existing = [int(x) for x in re.findall(r'<numFmt numFmtId="(\d+)"', mp)]
    base = max([167] + existing) + 1                        # id numFmt moi: tren moi id custom dang dung
    nm, new_nf = {}, []
    for i, el in enumerate(re.findall(r'<numFmt\b[^>]*/>', _inner(ps, "numFmts") or "")):
        pid = int(re.search(r'numFmtId="(\d+)"', el).group(1)); nid = base + i; nm[pid] = nid
        new_nf.append(re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % nid, el, count=1))
    add_csx = "".join(_remap_xf(x, nm, foff, filloff, boff)       for x in _xf_list(_inner(ps, "cellStyleXfs")))
    add_cx  = "".join(_remap_xf(x, nm, foff, filloff, boff, soff) for x in _xf_list(_inner(ps, "cellXfs")))
    mp = _add_container(mp, "numFmts", "".join(new_nf), len(new_nf), before_tag="fonts")
    mp = _add_container(mp, "fonts",   _inner(ps, "fonts")   or "", _count(ps, "fonts"))
    mp = _add_container(mp, "fills",   _inner(ps, "fills")   or "", _count(ps, "fills"))
    mp = _add_container(mp, "borders", _inner(ps, "borders") or "", _count(ps, "borders"))
    mp = _add_container(mp, "cellStyleXfs", add_csx, _count(ps, "cellStyleXfs"))
    mp = _add_container(mp, "cellXfs",      add_cx,  _count(ps, "cellXfs"))
    return mp, xoff
def _shift_sheet_styles(sheet, xoff):
    def add(m): return m.group(1) + str(int(m.group(2)) + xoff) + m.group(3)
    sheet = re.sub(r'(<c\b[^>]*?\bs=")(\d+)(")',     add, sheet)      # style o tung cell
    sheet = re.sub(r'(<row\b[^>]*?\bs=")(\d+)(")',   add, sheet)      # style ca dong
    sheet = re.sub(r'(<col\b[^>]*?\bstyle=")(\d+)(")', add, sheet)    # style cot (neu co)
    return sheet
def graft_pias(mapped_path, pias_bytes):
    """Ghep noi dung PIAS (1 file xlsx trong bo nho) thanh sheet 'Pias' cua file MAPPED."""
    zp = zipfile.ZipFile(io.BytesIO(pias_bytes))
    p_sheet  = zp.read("xl/worksheets/sheet1.xml").decode("utf-8")
    p_styles = zp.read("xl/styles.xml").decode("utf-8"); zp.close()
    try:
        zin = zipfile.ZipFile(mapped_path)
    except Exception as e:
        die(f"Cannot reopen '{os.path.basename(mapped_path)}' to add the Pias sheet. ({e})")
    names = zin.namelist(); parts = {n: zin.read(n) for n in names}; zin.close()
    # 1) gop styles
    merged, xoff = merge_styles(parts["xl/styles.xml"].decode("utf-8"), p_styles)
    parts["xl/styles.xml"] = merged.encode("utf-8")
    # 2) sheet PIAS: doi chi so style; bo cac phan tu phu thuoc part khac (CF/dxf, drawing, bang)
    p_sheet = _shift_sheet_styles(p_sheet, xoff)
    p_sheet = re.sub(r'<conditionalFormatting\b.*?</conditionalFormatting>', '', p_sheet, flags=re.S)
    p_sheet = re.sub(r'<tableParts\b.*?</tableParts>', '', p_sheet, flags=re.S)
    p_sheet = re.sub(r'<drawing\b[^>]*/>', '', p_sheet)
    p_sheet = re.sub(r'\btabSelected="1"', 'tabSelected="0"', p_sheet)   # Pias NOT selected -> avoid grouped sheets (grouping disables the AutoFilter dropdowns)
    newnum = max(int(re.search(r'sheet(\d+)\.xml', n).group(1))
                 for n in names if re.match(r'xl/worksheets/sheet\d+\.xml$', n)) + 1
    spart = "xl/worksheets/sheet%d.xml" % newnum
    parts[spart] = p_sheet.encode("utf-8"); names.append(spart)
    # 3) dang ky sheet vao workbook.xml (+ bat fullCalcOnLoad cho SUBTOTAL tu tinh)
    wb = parts["xl/workbook.xml"].decode("utf-8")
    new_sid = max(int(x) for x in re.findall(r'sheetId="(\d+)"', wb)) + 1
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    new_rid = "rId%d" % (max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels)) + 1)
    wb = wb.replace("</sheets>", '<sheet name="%s" sheetId="%d" r:id="%s"/></sheets>'
                                 % (PIAS_SHEET_NAME, new_sid, new_rid), 1)
    if re.search(r'<calcPr\b', wb):
        if 'fullCalcOnLoad' in wb:
            wb = re.sub(r'fullCalcOnLoad="\d+"', 'fullCalcOnLoad="1"', wb, count=1)
        else:
            wb = re.sub(r'(<calcPr\b[^>]*?)\s*/>', r'\1 fullCalcOnLoad="1"/>', wb, count=1)
    else:
        wb = wb.replace("</workbook>", '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>', 1)
    parts["xl/workbook.xml"] = wb.encode("utf-8")
    # 4) relationship + content type cho part sheet moi
    rels = rels.replace("</Relationships>",
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/></Relationships>'
        % (new_rid, newnum), 1)
    parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    ct = parts["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("</Types>",
        '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
        % newnum, 1)
    parts["[Content_Types].xml"] = ct.encode("utf-8")
    # 5) ghi lai file MAPPED (1 file duy nhat)
    tmp = mapped_path + ".tmp"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for n in names: zo.writestr(n, parts[n])
        os.replace(tmp, mapped_path)
    except PermissionError:
        if os.path.exists(tmp): os.remove(tmp)
        die(f"Cannot write '{os.path.basename(mapped_path)}' (is it open in Excel?). Close it and run again.")

# ---------- main ----------
def main():
    global GUI
    # Interactive run (double-click / drag-drop) + PowerShell available -> use dialog boxes
    # and hide the black console. Scripted run (sheet passed as 2nd arg) stays on the console.
    if len(sys.argv) <= 2 and _have_powershell():
        GUI = True; _hide_console()
    print("="*64)
    in_path = pick_input()
    print("INPUT :", os.path.basename(in_path))
    pick = choose_mapping(in_path)
    sheet, kind, hr = pick["sheet"], pick["kind"], pick["hr"]

    if kind == "data":
        claims, hr = read_data(in_path, sheet)
    else:
        rows = read_bangke(in_path, sheet, hr)
        claims = [dict(sheet_row=r["sheet_row"], no=None, rr=r["gsc"], serial=r["serial"],
                       phone=r["phone"], phone9=last9(r["phone"]), mgd=r["mgd"], si=r["si"],
                       incident=r["incident"], notif=r["notif"], cause=r["cause"],
                       detail=None, cost=r["cost"],
                       eff_file=r["eff_file"], exp_file=r["exp_file"], prod=r.get("prod")) for r in rows]
    print(f"SHEET : '{sheet}'  ({'Data (EN)' if kind == 'data' else 'bang ke'}, {len(claims)} rows)")
    if not claims:
        die("The chosen sheet has no claim rows to map.")
    group, family, plabel = choose_product()

    print("Connecting to Fabric and mapping...")
    _pg = ui_progress_start("Connecting to Fabric and mapping,\nplease wait...")
    cn = connect(); cur = cn.cursor()
    results = match(cur, claims, group, family)
    cn.close()
    ui_progress_stop(_pg)

    edits = [(sheet, hr, {m["c"]["sheet_row"]: _row_values(m) for m in results})]
    out_path = surgical_write(in_path, edits)
    pias_n = None
    if kind == "bangke":
        pb = gen_pias(out_path, sheet, hr)                # (bytes, n_rows) or None
        if pb:
            graft_pias(out_path, pb[0]); pias_n = pb[1]

    from collections import Counter
    matched = sum(1 for m in results if m["best"])
    bykey = Counter(m["key"] for m in results if m["best"])
    si_mis = sum(1 for m in results if m["best"] and m["c"]["si"] is not None
                 and pnum(m["best"]["SI"]) is not None and abs(pnum(m["best"]["SI"]) - m["c"]["si"]) >= 1)
    flagged = sum(1 for m in results if _row_values(m)["hsbt"])
    recreated = sum(1 for m in results if "đã hủy" in (m.get("note_extra") or ""))
    print("\n" + "="*64)
    print(f"  Sheet '{sheet}' : {len(results)} rows, mapped {matched} ({dict(bykey)}), SI mismatch {si_mis}")
    print(f"  Batch product   : {plabel}")
    print(f"  Rows with note  : {flagged}   (recreated: {recreated})")
    print(f"  OUTPUT          : {os.path.basename(out_path)}")
    if pias_n is not None:
        print(f"  + 'Pias' sheet  : {pias_n} rows (same file)")
    print("="*64)
    un = [m for m in results if not m["best"]]
    if un:
        print("  Rows NOT mapped (re-check serial/phone or tell Nam):")
        for m in un[:20]:
            c = m["c"]; print(f"   gsc/rr={c['rr']} serial={c['serial']!r} phone={c['phone']!r}")

    summary = (
        f"Sheet '{sheet}'  -  Product: {plabel}\n\n"
        f"Rows: {len(results)}    Mapped: {matched} ({dict(bykey)})\n"
        f"SI mismatch: {si_mis}    Not mapped: {len(un)}\n"
        f"Notes: {flagged}  (cancelled/recreated: {recreated})\n"
        + (f"Pias sheet: {pias_n} rows\n" if pias_n is not None else "")
        + f"\nOutput file (same folder as input):\n{os.path.basename(out_path)}"
        + (("\n\nNOT mapped (check serial/phone or tell Nam):\n" +
            "\n".join(f" - {m['c']['rr']}  serial {m['c']['serial']}" for m in un[:12])) if un else ""))
    if GUI:
        ui_msg("MAP CLAIM - Done", summary, "Information")
    else:
        try: input("\nPress Enter to close...")
        except Exception: pass

def signin_cli():
    """Refresh the Fabric sign-in on this machine (for the web host). Prints the device code
       in the console; the saved record is what app.py runs on. Usage: map_claim_ri.py --signin"""
    from azure.identity import (DeviceCodeCredential, TokenCachePersistenceOptions)
    def _p(u, c, e):
        print(f"\n  >>> Open in a browser: {u}\n  >>> Enter the code: {c}\n", flush=True)
    cred = DeviceCodeCredential(client_id=CLIENT,
        cache_persistence_options=TokenCachePersistenceOptions(name="claude_fabric_equip"),
        prompt_callback=_p)
    rec = cred.authenticate(scopes=["https://database.windows.net/.default"])
    open(AUTH, "w").write(rec.serialize())
    print("\n  Signed in as:", rec.username)
    print("  Saved to:", AUTH)
    cn = connect(); cn.close()
    print("  Fabric connection verified. The web app can query again.\n")

if __name__ == "__main__":
    try:
        if "--signin" in sys.argv:
            signin_cli(); sys.exit(0)
        main()
    except SystemExit:
        raise
    except BaseException as _err:
        _unhide_console()                       # never leave the user with a hidden console on a crash
        try: ui_msg("MAP CLAIM - Error", "Unexpected error:\n\n" + str(_err), "Error")
        except Exception: pass
        import traceback; traceback.print_exc()
        try: input("\nPress Enter to close...")
        except Exception: pass
        sys.exit(1)
