"""Check the built add-in the way Excel sees it, in private Excel instances. Nothing is saved or installed.

1. the package: customUI14.xml present, well-formed, wired from _rels/.rels; every onAction/onLoad names
   a Public Sub of Ribbon.bas;
2. opened in Excel: the VBA modules are the code of dist/*.bas|frm (the code the bench and the real
   books ran - compared case-insensitively, because the VBA editor re-cases identifiers across the
   whole project, e.g. `Cells` -> `cells`; string literals and comments must match exactly), the
   project compiles, the Add-ins dialog title/comments are set, every imageMso exists in this Excel;
3. always: a VISIBLE instance, moved off the screen, opens the example book and the add-in; the
   ribbon must load (Ribbon_OnLoad fired - Excel drops the WHOLE customization on any error in the
   ribbon XML, so this is the schema check by the real consumer) and the tab is shown;
   with --png PATH the window is also captured to PATH.
   An invisible automation instance builds no ribbon at all, so step 3 is the only ribbon check.

    python tests/verify_xlam.py dist [--png docs/ribbon.png]      exit 1 on any failure
"""
import sys, os, re, time, zipfile, xml.dom.minidom
import pythoncom, win32com.client

dist = os.path.abspath(sys.argv[1])
png = os.path.abspath(sys.argv[sys.argv.index('--png') + 1]) if '--png' in sys.argv else None
xlam = os.path.join(dist, 'excel-key-cleaner.xlam')
demo = os.path.join(os.path.dirname(dist), 'examples', 'dirty-keys-demo.xlsx')
failed = []
def check(ok, what):
    print(('ok    ' if ok else 'FAIL  ') + what)
    if not ok:
        failed.append(what)

# ---- 1. package ----
z = zipfile.ZipFile(xlam)
ui = z.read('customUI/customUI14.xml').decode('utf-8')
xml.dom.minidom.parseString(ui.encode('utf-8'))
check(True, 'customUI/customUI14.xml is well-formed XML')
import xml.etree.ElementTree as ET
REL_UI = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
rels = ET.fromstring(z.read('_rels/.rels')).findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
check(any(r.get('Type') == REL_UI and r.get('Target') == 'customUI/customUI14.xml' for r in rels),
      '_rels/.rels has the customUI14 relationship (exact Type and Target)')
ribbon_src = open(os.path.join(dist, 'Ribbon.bas'), encoding='cp1251').read()
# the signatures Office calls back with: onLoad(IRibbonUI), button onAction(IRibbonControl)
sigs = dict(re.findall(r'^Public Sub (\w+)\(\w+ As (IRibbonUI|IRibbonControl)\)', ribbon_src, re.M))
for kind, want in (('onLoad', 'IRibbonUI'), ('onAction', 'IRibbonControl')):
    for cb in sorted(set(re.findall(kind + r'="(\w+)"', ui))):
        check(sigs.get(cb) == want, f'{kind} {cb} is Public Sub {cb}(... As {want}) in Ribbon.bas')
icons = sorted(set(re.findall(r'imageMso="(\w+)"', ui)))

# ---- 2. in Excel ----
def module_lines(path):
    lines = open(path, encoding='cp1251').read().replace('\r\n', '\n').split('\n')
    if path.endswith('.frm'):          # the designer header is not part of the code module
        lines = lines[next(i for i, l in enumerate(lines) if l.startswith('Attribute VB_Name')):]
    return '\n'.join(l for l in lines if not l.startswith('Attribute ')).strip('\n').split('\n')

def split_line(line):
    """(code outside literals and comments, [string literals], comment) of one VBA line"""
    if re.match(r'\s*Rem(\s|$)', line, re.I):          # a Rem comment is a comment: compared exactly
        return '', [], line
    code, lits, cur, in_str, i = [], [], [], False, 0
    while i < len(line):
        ch = line[i]
        if in_str:
            if ch == '"' and line[i + 1:i + 2] == '"':
                cur.append('""'); i += 2; continue
            if ch == '"':
                lits.append(''.join(cur)); cur = []; in_str = False; code.append('"')
            else:
                cur.append(ch)
        elif ch == '"':
            in_str = True; code.append('"')
        elif ch == "'":
            return ''.join(code), lits, line[i:]
        else:
            code.append(ch)
        i += 1
    return ''.join(code), lits, ''

def same_code(want, got):
    if len(want) != len(got):
        return False, f'{len(want)} vs {len(got)} lines'
    for n, (a, b) in enumerate(zip(want, got), 1):
        if a == b:
            continue
        ca, la, ma = split_line(a)
        cb, lb, mb = split_line(b)
        if ca.lower() != cb.lower() or la != lb or ma != mb:
            return False, f'line {n}: {a!r} vs {b!r}'
    return True, ''

ICON_PROBE = '''
Public Function IconOk(ByVal name As String) As Boolean
    Dim p As Object
    On Error Resume Next
    Set p = Application.CommandBars.GetImageMso(name, 32, 32)
    IconOk = (Err.Number = 0 And Not p Is Nothing)
End Function
'''
pythoncom.CoInitialize()
xl = win32com.client.DispatchEx('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
wb = probe = None
try:
    probe = xl.Workbooks.Add()
    probe.VBProject.VBComponents.Add(1).CodeModule.AddFromString(ICON_PROBE)
    for name in icons:
        check(bool(xl.Run(f"'{probe.Name}'!IconOk", name)), f'imageMso {name} exists')
    wb = xl.Workbooks.Open(xlam)
    check(bool(wb.IsAddin), 'opens as an add-in (IsAddin)')
    check(wb.BuiltinDocumentProperties('Title').Value == 'Очистка ключей', 'title for the Add-ins dialog')
    check(len(str(wb.BuiltinDocumentProperties('Comments').Value)) > 20, 'comments for the Add-ins dialog')
    comps = {wb.VBProject.VBComponents.Item(i).Name: wb.VBProject.VBComponents.Item(i)
             for i in range(1, wb.VBProject.VBComponents.Count + 1)}
    for mod, fname in (('Cleaning', 'Cleaning.bas'), ('CleaningForm', 'CleaningForm.frm'), ('Ribbon', 'Ribbon.bas')):
        if mod not in comps:
            check(False, f'module {mod} present'); continue
        cm = comps[mod].CodeModule
        ok, why = same_code(module_lines(os.path.join(dist, fname)),
                            cm.Lines(1, cm.CountOfLines).replace('\r\n', '\n').strip('\n').split('\n'))
        check(ok, f'module {mod} == dist/{fname} (code case-insensitive, strings and comments exact) {why}')
    xl.Run(f"'{wb.Name}'!Ribbon.RibbonLoaded")        # a compile error would surface as "cannot run the macro"
    check(True, 'project compiles')
    # the form's designer layer (.frx): the controls with their captions, positions, sizes and tab order must be
    # what a plain import of dist/CleaningForm.frm gives. Not compared as bytes: two import->export round trips
    # of the same .frm in the same Excel already differ inside the .frx streams (Excel rewrites counters).
    def form_signature(comp):
        d = comp.Designer
        sig = [(str(comp.Properties('Caption').Value), comp.Properties('Width').Value, comp.Properties('Height').Value)]
        for i in range(d.Controls.Count):
            c = d.Controls.Item(i)
            sig.append((c.Name, getattr(c, 'Caption', None), c.Left, c.Top, c.Width, c.Height, c.TabIndex))
        return sig
    fresh = xl.Workbooks.Add()
    try:
        fresh.VBProject.VBComponents.Import(os.path.join(dist, 'CleaningForm.frm'))
        want_form = form_signature(fresh.VBProject.VBComponents('CleaningForm'))
        got_form = form_signature(comps['CleaningForm'])
        check(got_form == want_form, f'form CleaningForm: {len(got_form) - 1} controls, captions, positions, tab order == a plain import of dist/')
    except Exception as e:                       # e.g. dist/CleaningForm.frm did not import as a form at all
        check(False, f'form CleaningForm comparable with a plain import of dist/: {e}')
    finally:
        fresh.Close(False)
finally:
    for b in (wb, probe):
        if b is not None:
            b.Close(False)
    xl.Quit()

# ---- 3. the ribbon, in a visible instance moved off the screen ----
import ctypes
from ctypes import wintypes
from PIL import Image
user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32

def capture(hwnd, path, height):
    r = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    w, h = r.right - r.left, r.bottom - r.top
    hdc = user32.GetWindowDC(hwnd); mdc = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, w, h); gdi32.SelectObject(mdc, bmp)
    ok = user32.PrintWindow(hwnd, mdc, 2)          # PW_RENDERFULLCONTENT: works off the screen
    class BIH(ctypes.Structure):
        _fields_ = [(n, t) for n, t in (('biSize', ctypes.c_uint32), ('biWidth', ctypes.c_int32),
                    ('biHeight', ctypes.c_int32), ('biPlanes', ctypes.c_uint16), ('biBitCount', ctypes.c_uint16),
                    ('biCompression', ctypes.c_uint32), ('biSizeImage', ctypes.c_uint32),
                    ('biXPelsPerMeter', ctypes.c_int32), ('biYPelsPerMeter', ctypes.c_int32),
                    ('biClrUsed', ctypes.c_uint32), ('biClrImportant', ctypes.c_uint32))]
    bih = BIH(ctypes.sizeof(BIH), w, -h, 1, 32, 0, 0, 0, 0, 0, 0)
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(mdc, bmp, 0, h, buf, ctypes.byref(bih), 0)
    img = Image.frombuffer('RGBA', (w, h), buf, 'raw', 'BGRA', 0, 1).convert('RGB')
    img.crop((0, 0, w, min(h, height))).save(path)
    gdi32.DeleteObject(bmp); gdi32.DeleteDC(mdc); user32.ReleaseDC(hwnd, hdc)
    return bool(ok)

xl = win32com.client.DispatchEx('Excel.Application')
xl.DisplayAlerts = False
book = ad = None
try:
    xl.Visible = True
    xl.WindowState = -4143                          # xlNormal
    xl.Width, xl.Height = 1100, 520
    xl.Left, xl.Top = -3000, 0                     # off the user's screen
    book = xl.Workbooks.Open(demo, 0, True)
    ad = xl.Workbooks.Open(xlam)
    for _ in range(20):                             # the ribbon is built asynchronously
        if xl.Run(f"'{ad.Name}'!Ribbon.RibbonLoaded"):
            break
        time.sleep(0.25)
    loaded = bool(xl.Run(f"'{ad.Name}'!Ribbon.RibbonLoaded"))
    check(loaded, 'ribbon loaded in a visible instance (Ribbon_OnLoad fired)')
    if loaded:
        xl.Run(f"'{ad.Name}'!Ribbon.ShowTab")
        time.sleep(1.0)
        if png:
            check(capture(xl.Hwnd, png, 460), f'ribbon captured to {png}')
finally:
    for b in (ad, book):
        if b is not None:
            b.Close(False)
    xl.Quit()

print('\nALL OK' if not failed else f'\n{len(failed)} FAILED')
sys.exit(1 if failed else 0)
