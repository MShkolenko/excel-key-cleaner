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
FAIL_AT = ''
FAIL_POST = False
FAIL_RESTORE = False
FAIL_CLOSE = False
LONG_ERR = False
for _a in sys.argv[3:]:
    if _a.startswith('--fail-at='): FAIL_AT = _a.split('=', 1)[1]
    if _a == '--fail-post': FAIL_POST = True
    if _a == '--fail-restore': FAIL_RESTORE = True
    if _a == '--fail-all': FAIL_POST = FAIL_RESTORE = FAIL_CLOSE = LONG_ERR = True
OUT_NAME = 'Cleaning_Test.bas'
if FAIL_AT: OUT_NAME = 'Cleaning_TestFail.bas'
if FAIL_POST: OUT_NAME = 'Cleaning_TestPostFail.bas'
if FAIL_RESTORE: OUT_NAME = 'Cleaning_TestRestoreFail.bas'
if FAIL_CLOSE: OUT_NAME = 'Cleaning_TestAllFail.bas'

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
    elif FAIL_RESTORE and ln.strip().startswith('Err.Clear: Application.EnableEvents = ee'):
        # fault injection: in the TEST copy the restore silently does not take, so the read-back
        # disagrees - exactly the case RestoreApp must report instead of swallowing
        ln = "    Err.Clear: got = Not ee   'injected test failure: the restore did not take"
    elif FAIL_RESTORE and ln.strip().startswith('Err.Clear: Application.EnableEvents = ee'):
        # fault injection: in the TEST copy the restore silently does not take, so the read-back
        # disagrees - exactly the case RestoreApp must report instead of swallowing
        ln = "    Err.Clear: got = Not ee   'injected test failure: the restore did not take"
    elif FAIL_CLOSE and ln.strip() == 'wb.Close SaveChanges:=False':
        # fault injection: the backup workbook refuses to close (Err.Number <> 0 -> CloseBackup False)
        ln = '    Err.Raise 1004, "CloseBackup", "injected backup close failure"'
    elif FAIL_POST and ln.startswith('Private Function SelectVerified('):
        # fault injection for the post-write probe: the TEST copy fails AFTER the last write, in the
        # region the new PostWrite handler covers (nothing to roll back, everything to clean up)
        _desc = '"injected test failure after the writes" & String$(1200, "z")' if LONG_ERR else '"injected test failure after the writes"'
        ln = ln + '\r\n    Err.Raise 1004, "SelectVerified", ' + _desc
    elif FAIL_AT and ln.startswith('Private Sub PutText('):
        # fault injection for the rollback probe: the TEST copy fails to write one given address
        ln = ln + '\r\n    If c.Address(False, False) = "' + FAIL_AT + '" Then Err.Raise 1004, "PutText", "injected test failure at ' + FAIL_AT + '"'
    new.append(ln)
harness = r'''
'===================== test harness (not part of the macro) =====================
Private Function MsgBox(Prompt As Variant, Optional Buttons As Variant, Optional Title As Variant) As Long
    If Len(LastMsg) > 0 Then LastMsg = LastMsg & vbCrLf & "|NEXT MSGBOX|" & vbCrLf
    LastMsg = LastMsg & CStr(Prompt)
End Function

Public Function RunTest(c1 As Boolean, c2 As Boolean, c3 As Boolean, c4 As Boolean, c5 As Boolean) As String
    On Error GoTo EH
    LastMsg = ""
    CleaningForm.CheckBox1.Value = c1
    CleaningForm.CheckBox2.Value = c2
    CleaningForm.CheckBox3.Value = c3
    CleaningForm.CheckBox4.Value = c4
    CleaningForm.Controls("CheckBox5").Value = c5
    CleanKeys
    RunTest = LastMsg & "|SELECTED=" & Selection.Address(False, False)
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
(HERE / OUT_NAME).write_bytes(test_src.encode('cp1251'))

diff = list(difflib.unified_diff(lines, new, 'Cleaning.bas', 'Cleaning_Test.bas (before harness)', lineterm='', n=0))
print('--- diff of the macro body (must be exactly 2 hunks: VB_Name and the .Show stub) ---')
print('\n'.join(diff))

# ---------- 2. cases ----------
ESC = {'{LF}': '\n', '{CR}': '\r', '{TAB}': '\t', '{NBSP}': '\u00a0', '{ZWSP}': '\u200b', '{ZWNJ}': '\u200c',
       '{ZWJ}': '\u200d', '{BOM}': '\ufeff', '{ENSP}': '\u2002', '{THIN}': '\u2009', '{IDEO}': '\u3000', '{WJ}': '\u2060',
       '{LSEP}': '\u2028', '{PSEP}': '\u2029', '{OGH}': '\u1680',
       '{NEL}': '\u0085', '{VT}': '\u000b', '{FF}': '\u000c'}
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
    ('c32', 'a{NEL}b{VT}c{FF}d', 'NEL, vertical tab, form feed are line breaks too (Codex round 2)'),
    ('c33', 'X{CR}{LF} Y', 'CRLF followed by a real space: one space must survive as the original'),
    ('c34', '{NBSP}{ZWSP}AB1Р ', 'edge junk with ONLY the Cyrillic action on - trimmed anyway'),
    ('c35', '{LF}AB ', 'a line break at the edge stays unless action 2 is on; the space always goes'),
    ('c36', '{LF} AB', 'a KEPT line break must not shield the space next to it'),
    ('c37', 'AB {LF}', 'the same at the trailing edge'),
    ('c38', '{LF}  ', 'the whole cell is whitespace AND a break must be kept - one run, not two'),
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

def mirror2(s, c, l, r, n, o=0):
    t = s
    if c:
        cyr = [ch for ch in t if 'Ѐ' <= ch <= 'ӿ']
        n_latin = sum(1 for ch in t if ('A' <= ch <= 'Z') or ('a' <= ch <= 'z'))
        has_digit = any('0' <= ch <= '9' for ch in t)
        # code-like: a digit, a Latin letter, at least as many Latin as Cyrillic letters, all Cyrillic are homoglyphs
        if cyr and (o or (has_digit and n_latin >= len(cyr) and all(ch in MAP2 for ch in cyr))):
            t = ''.join(MAP2.get(ch, ch) for ch in t)
    if l:
        t = t.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ').replace('\u2028', ' ').replace('\u2029', ' ').replace('\u0085', ' ').replace('\u000b', ' ').replace('\u000c', ' ')
    if r:
        for ch in WS2: t = t.replace(ch, ' ')
        for ch in ZW2: t = t.replace(ch, '')
        t = t.replace(' ', '')
    if n:
        for ch in WS2: t = t.replace(ch, ' ')
        for ch in ZW2: t = t.replace(ch, '')
        while '  ' in t: t = t.replace('  ', ' ')
        t = t.strip(' ')
    # edges are trimmed ALWAYS (operator, 2026-09-22), whatever the checkboxes say - including none
    # of them. A line break at the edge survives only while action 2 is off, and it does NOT shield
    # the ordinary spaces around it (Codex round 7).
    BREAKS = '\r\n\u2028\u2029\u0085\u000b\u000c'
    space_like = WS2 + ZW2 + ' ' + BREAKS
    head = len(t) - len(t.lstrip(space_like))
    tail = 0 if head == len(t) else len(t) - len(t.rstrip(space_like))   # all-whitespace: ONE run, not two
                                                                          # (Codex round 8: the kept break was counted twice)
    keep_head = '' if l else ''.join(ch for ch in t[:head] if ch in BREAKS)
    keep_tail = '' if l else ''.join(ch for ch in t[len(t) - tail:] if ch in BREAKS)
    t = keep_head + t[head:len(t) - tail] + keep_tail
    return t

def mirror(s, c, l, r, n, o=0):
    if V2: return mirror2(s, c, l, r, n, o)
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
    'D':  (1, 1, 0, 1, 0),   # form defaults
    'O':  (1, 1, 0, 1, 1),   # defaults + 'selection is codes only': Cyrillic converted everywhere
    'C':  (1, 0, 0, 0, 0),
    'L':  (0, 1, 0, 0, 0),
    'R':  (0, 0, 1, 0, 0),
    'N':  (0, 0, 0, 1, 0),
    'LR': (0, 1, 1, 0, 0),
    'Z':  (0, 0, 0, 0, 0),   # nothing checked at all: only the always-on edge cleaning may happen
}
with open(HERE / 'opts.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['opt', 'c1', 'c2', 'c3', 'c4', 'c5'])
    for k, v in OPTS.items():
        w.writerow([k] + [int(x) for x in v])

with open(HERE / 'expected.csv', 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(['opt', 'id', 'expected', 'cyr_left'])
    for k, v in OPTS.items():
        for cid, inp, _ in CASES:
            out = mirror(unesc(inp), *v)
            cyr_left = int(bool(v[0]) and any('\u0400' <= ch <= '\u04ff' for ch in out))
            w.writerow([k, cid, esc(out), cyr_left])
print(f'\n{len(CASES)} cases x {len(OPTS)} option sets -> cases.csv, opts.csv, expected.csv, Cleaning_Test.bas')
