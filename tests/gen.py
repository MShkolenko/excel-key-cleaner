"""Build the headless test module + case/expectation tables for the Cleaning macro audit.

Usage:  python gen.py <Cleaning.bas exported by Export-VbaComponentsNative.ps1> <out dir> [--v2]
        powershell -File Run-CleaningAudit.ps1 -AuditDir <out dir> -FormDir <dir with CleaningForm.frm/.frx>
        python verify.py <out dir>
The .bas must be the native export (cp1251, with the Attribute VB_Name line) - not Export-VBA.ps1's text dump.
--v2 selects the mirror of the reworked macro (code-only homoglyphs, all whitespace kinds, trim).
"""
import sys, csv, difflib, pathlib
sys.stdout.reconfigure(encoding='utf-8')

SRC = pathlib.Path(sys.argv[1])
HERE = pathlib.Path(sys.argv[2])
HERE.mkdir(parents=True, exist_ok=True)

# ---------- 1. test module: verbatim copy, form call stubbed, MsgBox shadowed ----------
orig = SRC.read_bytes().decode('cp1251')
lines = orig.splitlines()
new = []
for ln in lines:
    if ln == 'Attribute VB_Name = "Cleaning"':
        ln = 'Attribute VB_Name = "Cleaning_Test"'
        if 'Option Explicit' not in orig:
            ln += '\r\nPublic LastMsg As String   ' + "'test harness"
    elif ln.strip() == 'Option Explicit':
        ln = 'Option Explicit\r\nPublic LastMsg As String   ' + "'test harness"
    elif ln.strip() == 'CleaningForm.Show':
        # works for v1 (which only checks for "Cancel") and v2 (which requires "OK")
        ln = '    CleaningForm.Tag = "OK"   ' + "'stubbed CleaningForm.Show: the headless test sets the checkboxes directly"
    new.append(ln)
harness = r'''
'===================== test harness (not part of the macro) =====================
Private Function MsgBox(Prompt As Variant, Optional Buttons As Variant, Optional Title As Variant) As Long
    LastMsg = CStr(Prompt)
End Function

Public Function RunTest(c1 As Boolean, c2 As Boolean, c3 As Boolean, c4 As Boolean) As String
    On Error GoTo EH
    LastMsg = ""
    CleaningForm.CheckBox1.Value = c1
    CleaningForm.CheckBox2.Value = c2
    CleaningForm.CheckBox3.Value = c3
    CleaningForm.CheckBox4.Value = c4
    Замена_Кирилицы_С_Выбором
    RunTest = LastMsg
    Exit Function
EH:
    RunTest = "VBA ERROR " & Err.Number & ": " & Err.Description & " | ScreenUpdating=" & Application.ScreenUpdating
End Function

' What the form's close box (X) leads to: default QueryClose unloads the instance, and the
' macro's next reference to CleaningForm silently creates a fresh one.
Public Function ProbeCloseX() As String
    CleaningForm.Tag = "Cancel"
    CleaningForm.CheckBox1.Value = False
    Unload CleaningForm
    ProbeCloseX = "after Unload: Tag=[" & CleaningForm.Tag & "] CheckBox1=" & CleaningForm.CheckBox1.Value & " CheckBox2=" & CleaningForm.CheckBox2.Value & " CheckBox4=" & CleaningForm.CheckBox4.Value
End Function

' Codex's proposed fix: force text format around the write, then restore the old format.
Public Function ProbeSafeWrite(c As Range, v As String) As String
    Dim oldFmt As Variant
    oldFmt = c.NumberFormat
    c.NumberFormat = "@"
    c.Value2 = v
    c.NumberFormat = oldFmt
    ProbeSafeWrite = "value=[" & c.Value2 & "] type=" & TypeName(c.Value2) & " formula=" & c.HasFormula & " fmt=[" & c.NumberFormat & "]"
End Function
'''
test_src = '\r\n'.join(new) + '\r\n' + harness.replace('\n', '\r\n')
(HERE / 'Cleaning_Test.bas').write_bytes(test_src.encode('cp1251'))

diff = list(difflib.unified_diff(lines, new, 'Cleaning.bas', 'Cleaning_Test.bas (before harness)', lineterm='', n=0))
print('--- diff of the macro body (must be exactly 2 hunks: VB_Name and the .Show stub) ---')
print('\n'.join(diff))

# ---------- 2. cases ----------
ESC = {'{LF}': '\n', '{CR}': '\r', '{TAB}': '\t', '{NBSP}': '\u00a0', '{ZWSP}': '\u200b', '{ZWNJ}': '\u200c',
       '{ZWJ}': '\u200d', '{BOM}': '\ufeff', '{ENSP}': '\u2002', '{THIN}': '\u2009', '{IDEO}': '\u3000', '{WJ}': '\u2060',
       '{LSEP}': '\u2028', '{PSEP}': '\u2029', '{OGH}': '\u1680'}
def unesc(s):
    for k, v in ESC.items():
        s = s.replace(k, v)
    return s
def esc(s):
    for k, v in ESC.items():
        s = s.replace(v, k)
    return s

CASES = [
    ('c01', 'XX.U1.ABC.0001.10UРH.0.ZZ.YY0001-ST00', 'Cyrillic Р inside a long structured key (fictional)'),
    ('c02', '42UРQ', 'Cyrillic Р inside a short code (fictional)'),
    ('c03', 'ВЕТЕР', 'pure Russian word - every letter is a homoglyph'),
    ('c04', 'СМР работы', 'Russian prose with lowercase'),
    ('c05', 'cе.u1.rpr', 'lowercase Cyrillic е'),
    ('c06', 'УКС', 'Cyrillic У (looks like Y) is not in the map'),
    ('c07', 'A  B   C', 'runs of spaces'),
    ('c08', ' ABC ', 'leading/trailing space'),
    ('c09', 'A{NBSP}B', 'non-breaking space'),
    ('c10', 'A{TAB}B', 'tab'),
    ('c11', 'line1{LF}line2', 'LF'),
    ('c12', 'line1{CR}{LF}line2', 'CRLF'),
    ('c13', 'line1{CR}line2', 'CR only'),
    ('c14', '   ', 'spaces only'),
    ('c15', 'ЁЖ', 'Cyrillic with no homoglyph'),
    ('c16', '  ХХ.U1 {LF} ABC  ', 'everything at once'),
    ('c17', 'ABC', 'pure Latin - must be untouched'),
    ('c18', 'ХОРОШО', 'Russian word, partly homoglyphs'),
    ('c19', 'Р', 'single Cyrillic char'),
    ('c20', 'A{NBSP}{NBSP}B  C', 'nbsp run next to space run'),
    ('c21', 'XX.U1{ZWSP}.ABC{ZWNJ}{ZWJ}{BOM}', 'zero-width chars inside a key'),
    ('c22', 'A{ENSP}B{THIN}C{IDEO}D{WJ}E', 'exotic spaces and word joiner'),
    ('c23', '07UУQ', 'Cyrillic У inside a Latin code: NOT converted (У/у deliberately not mapped), counted as ambiguous'),
    ('c24', '{ZWSP}', 'only a zero-width space'),
    ('c25', 'a{LSEP}b{PSEP}c', 'Unicode line/paragraph separators'),
    ('c26', 'a{OGH}b', 'Ogham space mark'),
    ('c27', 'МОСТ-A', 'Russian word + Latin suffix, no digit: must stay (Codex v2.1 finding)'),
    ('c28', 'ВЕТЕР A', 'Russian word + Latin letter, no digit: must stay'),
    ('c29', 'СОРТ-A1', 'Russian word of homoglyphs + digit, but Latin < Cyrillic: must stay'),
    ('c30', 'СЕ.U2.ABC.0002', 'whole segment typed in Cyrillic inside a Latin key: fixed'),
    ('c31', 'СЕ.U1', 'short key, Latin < Cyrillic: left alone and counted (documented limit)'),
]
with open(HERE / 'cases.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f, quoting=csv.QUOTE_ALL)
    w.writerow(['id', 'input', 'note'])
    for cid, inp, note in CASES:
        w.writerow([cid, inp, note])

# ---------- 3. mirror of the VBA semantics ----------
V2 = '--v2' in sys.argv
MAP = dict(zip('АВЕКМНОРСТХ', 'ABEKMHOPCTX'))
MAP2 = dict(zip('АВЕКМНОРСТХаеорсх', 'ABEKMHOPCTXaeopcx'))   # no У/у: operator decision 2026-09-21
WS2 = '\t\u00a0\u1680' + ''.join(chr(c) for c in range(0x2000, 0x200b)) + '\u202f\u205f\u3000'
ZW2 = '\u200b\u200c\u200d\u2060\ufeff'   # zero-width: deleted, not turned into a space

def mirror2(s, c, l, r, n):
    t = s
    if c:
        cyr = [ch for ch in t if 'Ѐ' <= ch <= 'ӿ']
        n_latin = sum(1 for ch in t if ('A' <= ch <= 'Z') or ('a' <= ch <= 'z'))
        has_digit = any('0' <= ch <= '9' for ch in t)
        # code-like: a digit, a Latin letter, at least as many Latin as Cyrillic letters, all Cyrillic are homoglyphs
        if cyr and has_digit and n_latin >= len(cyr) and all(ch in MAP2 for ch in cyr):
            t = ''.join(MAP2.get(ch, ch) for ch in t)
    if l:
        t = t.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ').replace('\u2028', ' ').replace('\u2029', ' ')
    if r:
        for ch in WS2: t = t.replace(ch, ' ')
        for ch in ZW2: t = t.replace(ch, '')
        t = t.replace(' ', '')
    if n:
        for ch in WS2: t = t.replace(ch, ' ')
        for ch in ZW2: t = t.replace(ch, '')
        while '  ' in t: t = t.replace('  ', ' ')
        t = t.strip(' ')
    return t

def mirror(s, c, l, r, n):
    if V2: return mirror2(s, c, l, r, n)
    t = s
    if c:
        t = ''.join(MAP.get(ch, ch) for ch in t)
    if l:
        t = t.replace('\n', ' ').replace('\r', ' ')
        while True:
            u = t.replace('  ', ' ')
            if u == t: break
            t = u
    if r:
        t = t.replace(' ', '')
    if n:
        while True:
            u = t.replace('    ', ' ').replace('   ', ' ').replace('  ', ' ')
            if u == t: break
            t = u
    return t

OPTS = {
    'D':  (1, 1, 0, 1),   # form defaults
    'C':  (1, 0, 0, 0),
    'L':  (0, 1, 0, 0),
    'R':  (0, 0, 1, 0),
    'N':  (0, 0, 0, 1),
    'LR': (0, 1, 1, 0),
}
with open(HERE / 'opts.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['opt', 'c1', 'c2', 'c3', 'c4'])
    for k, v in OPTS.items():
        w.writerow([k] + [int(x) for x in v])

with open(HERE / 'expected.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['opt', 'id', 'expected'])
    for k, v in OPTS.items():
        for cid, inp, _ in CASES:
            w.writerow([k, cid, esc(mirror(unesc(inp), *v))])
print(f'\n{len(CASES)} cases x {len(OPTS)} option sets -> cases.csv, opts.csv, expected.csv, Cleaning_Test.bas')
