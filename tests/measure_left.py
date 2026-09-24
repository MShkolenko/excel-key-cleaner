"""Speed + selection gate for the case a real sheet exposed (v3.1, 2026-09-24).

Half the cells are suspicious leftovers the macro must SELECT (since v3.3 a Latin and a Cyrillic
letter in one word, "СЕ.U7", which the rule does not fix), laid out so that no two of them
touch (two columns, alternating rows): every one becomes its own area of the final selection.
The one-cell-at-a-time Union that used to build that selection grows roughly as the CUBE of the
area count - a real 63 109-cell sheet took 15 minutes. tests/Measure-Speed.ps1 never hit this: its
cells are all code-like, so nothing is left to name.

Runs the REAL dist/Cleaning.bas (module renamed, form .Show stubbed, MsgBox shadowed - the same
three substitutions as gen.py) in a fresh invisible Excel. Nothing is saved. Exit code 1 when the
report says the selection failed, when fewer cells are selected than named, or when time grows
faster than N^1.5 between sizes.

    python tests/measure_left.py dist            [sizes...]   default: 2000 8000 20000
"""
import sys, os, time, tempfile
import pythoncom, win32com.client

dist = os.path.abspath(sys.argv[1])
sizes = [int(a) for a in sys.argv[2:]] or [2000, 8000, 20000]
LEFT = 'СЕ.U'          # + a number: Latin 1 < Cyrillic 2 -> left, and mixed -> selected
CYR_R = 'Р'

src = open(os.path.join(dist, 'Cleaning.bas'), encoding='cp1251').read()
src = src.replace('Attribute VB_Name = "Cleaning"', 'Attribute VB_Name = "Cleaning_Left"')
src = src.replace('CleaningForm.Show', 'CleaningForm.Tag = "OK"')
src = src.replace('Option Explicit', 'Option Explicit\r\nPublic LastMsg As String', 1)
src += '''
Private Function MsgBox(Prompt As Variant, Optional Buttons As Variant, Optional Title As Variant) As Long
    LastMsg = LastMsg & CStr(Prompt) & vbLf
End Function
Public Function RunNow() As String
    LastMsg = ""
    CleaningForm.CheckBox1.Value = True
    CleaningForm.CheckBox2.Value = True
    CleaningForm.CheckBox3.Value = False
    CleaningForm.CheckBox4.Value = True
    CleaningForm.Controls("CheckBox5").Value = False
    CleanKeys
    RunNow = LastMsg & "|SELECTED CELLS=" & Selection.CountLarge
End Function
'''.replace('\n', '\r\n')
tmp = os.path.join(tempfile.gettempdir(), 'Cleaning_Left.bas')
with open(tmp, 'w', encoding='cp1251', newline='') as f:
    f.write(src)

failed = 0
pythoncom.CoInitialize()
xl = win32com.client.DispatchEx('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Add()
    wb.VBProject.VBComponents.Import(tmp)
    wb.VBProject.VBComponents.Import(os.path.join(dist, 'CleaningForm.frm'))
    macro = f"'{wb.Name}'!Cleaning_Left.RunNow"
    prev = None
    print(f'{"N":>7} {"left":>7} {"seconds":>8}  verdict')
    for n in sizes:
        ws = wb.Worksheets.Add()
        rows = n // 2
        data = []
        for i in range(rows):
            code, word = f'AB{i} {CYR_R}', f'{LEFT}{i}'
            data.append((word, code) if i % 2 else (code, word))
        rng = ws.Range(ws.Cells(1, 2), ws.Cells(rows, 3))
        rng.NumberFormat = '@'
        rng.Value2 = data
        rng.Select()
        t = time.perf_counter()
        msg = str(xl.Run(macro))
        secs = time.perf_counter() - t
        selected = int(msg.rsplit('|SELECTED CELLS=', 1)[1])
        verdict = []
        if 'НЕ удалось' in msg:
            verdict.append('report says the selection FAILED')
        if selected < rows:
            verdict.append(f'selected {selected} < named {rows}')
        if prev and secs > 1.0 and secs / prev[1] > (n / prev[0]) ** 1.5:
            verdict.append(f'time x{secs / prev[1]:.1f} for N x{n / prev[0]:.1f} - worse than N^1.5')
        failed += bool(verdict)
        print(f'{n:>7} {rows:>7} {secs:>8.2f}  ' + ('; '.join(verdict) or 'OK'))
        prev = (n, secs)
finally:
    if wb is not None:
        wb.Close(False)
    xl.Quit()
sys.exit(1 if failed else 0)
