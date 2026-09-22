Attribute VB_Name = "Cleaning"
Option Explicit

' Очистка ключей в выделенных ячейках: кириллические омоглифы -> латиница, переносы строк,
' пробелы. Работает только с текстовыми константами (формулы, числа, даты, ошибки, пустые и
' скрытые ячейки пропускаются), пишет через формат "@" (текст остаётся текстом), считает план
' до первой записи и откатывает уже записанное при ошибке.

' У/у намеренно НЕ в списке: русская у и латинская Y - не одно и то же (оператор, 2026-09-21).
Private Const HOMO_FROM As String = "АВЕКМНОРСТХаеорсх"
Private Const HOMO_TO As String = "ABEKMHOPCTXaeopcx"

Private Type Stats
    cellsSeen As Long
    cellsChanged As Long
    homoglyphs As Long
    ambiguous As Long
    lineBreaks As Long
    spacesRemoved As Long
    spacesCollapsed As Long
    richText As Long
    hidden As Long
End Type

' Старое имя оставлено как псевдоним - на нём у людей кнопки.
Public Sub Замена_Кирилицы_С_Выбором()
    CleanKeys
End Sub

Public Sub CleanKeys()
    Dim doCyrillic As Boolean, doLineBreaks As Boolean, doRemoveSpaces As Boolean, doNormalizeSpaces As Boolean
    Dim target As Range, cell As Range
    Dim st As Stats
    Dim plannedCells As New Collection, plannedOld As New Collection, plannedNew As New Collection, plannedFmt As New Collection
    Dim s0 As String, s As String
    Dim i As Long, attempted As Long, rolledBack As Long, rollbackFailed As Long
    Dim oldSU As Boolean, oldEE As Boolean, oldCalc As XlCalculation, calcChanged As Boolean

    If TypeName(Selection) <> "Range" Then
        MsgBox "Выделите ячейки для обработки.", vbExclamation
        Exit Sub
    End If

    ' ---------- форма: продолжаем ТОЛЬКО по кнопке "Выполнить" ----------
    CleaningForm.Tag = vbNullString
    CleaningForm.Show
    If CleaningForm.Tag <> "OK" Then
        Unload CleaningForm
        Exit Sub
    End If
    doCyrillic = CleaningForm.CheckBox1.Value
    doLineBreaks = CleaningForm.CheckBox2.Value
    doRemoveSpaces = CleaningForm.CheckBox3.Value
    doNormalizeSpaces = CleaningForm.CheckBox4.Value
    Unload CleaningForm

    If Not (doCyrillic Or doLineBreaks Or doRemoveSpaces Or doNormalizeSpaces) Then
        MsgBox "Вы не выбрали ни одного действия.", vbExclamation, "Нет действий"
        Exit Sub
    End If

    ' ---------- только текстовые константы, только видимые ----------
    Set target = TextConstantsIn(Selection, st.hidden)
    If target Is Nothing Then
        MsgBox "В выделении нет видимых текстовых ячеек." & vbCrLf & _
               "Формулы, числа, даты, ошибки, пустые и скрытые ячейки не обрабатываются.", vbInformation
        Exit Sub
    End If

    ' ---------- план: что во что превратится, без единой записи ----------
    For Each cell In target
        st.cellsSeen = st.cellsSeen + 1
        If IsRichText(cell) Then
            st.richText = st.richText + 1
        Else
            s0 = cell.Value2
            s = s0
            If doCyrillic Then s = FixHomoglyphs(s, st)
            If doLineBreaks Then s = FixLineBreaks(s, st)
            If doRemoveSpaces Then s = RemoveSpaces(s, st)
            If doNormalizeSpaces Then s = NormalizeSpaces(s, st)
            If s <> s0 Then
                If cell.Parent.ProtectContents And cell.Locked Then
                    MsgBox "Ячейка " & cell.Address(False, False) & " заблокирована на защищённом листе." & vbCrLf & _
                           "Ничего не изменено.", vbCritical, "Защищённый лист"
                    Exit Sub
                End If
                plannedCells.Add cell
                plannedOld.Add s0
                plannedNew.Add s
                plannedFmt.Add cell.NumberFormat
            End If
        End If
    Next cell

    ' ---------- запись с откатом ----------
    ' Обработчик взводится ДО первого касания Application.*: даже смена режима пересчёта может
    ' упасть (книга с OLAP-источником допускает только автоматический режим), и тогда экран и
    ' события должны быть возвращены как были.
    attempted = 0
    On Error GoTo Fail
    oldSU = Application.ScreenUpdating
    oldEE = Application.EnableEvents
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    calcChanged = True

    For i = 1 To plannedCells.Count
        attempted = i
        PutText plannedCells(i), plannedNew(i)
    Next i
    st.cellsChanged = attempted
    On Error GoTo 0

    RestoreApp oldSU, oldEE, oldCalc, calcChanged
    MsgBox Report(st, doCyrillic, doLineBreaks, doRemoveSpaces, doNormalizeSpaces), vbInformation, "Результат обработки"
    Exit Sub

Fail:
    Dim errText As String
    errText = Err.Description
    ' Откат покрывает и ячейку, на которой упала запись: PutText - три операции, любая из них
    ' могла уже сработать. Возвращаются и значение, и прежний числовой формат.
    On Error Resume Next
    For i = attempted To 1 Step -1
        Err.Clear
        RestoreCell plannedCells(i), plannedOld(i), plannedFmt(i)
        If Err.Number = 0 Then rolledBack = rolledBack + 1 Else rollbackFailed = rollbackFailed + 1
    Next i
    RestoreApp oldSU, oldEE, oldCalc, calcChanged
    Dim msg As String
    msg = "Ошибка при записи: " & errText & vbCrLf & vbCrLf
    If attempted = 0 Then
        msg = msg & "Ни одна ячейка не изменена."
    ElseIf rollbackFailed = 0 Then
        msg = msg & "Затронутые ячейки (" & attempted & ") возвращены к исходным значениям - проверено поячеечно."
    Else
        msg = msg & "Возвращено ячеек: " & rolledBack & ", НЕ удалось вернуть: " & rollbackFailed & _
              " (первая из затронутых: " & plannedCells(1).Address(False, False) & ", последняя: " & _
              plannedCells(attempted).Address(False, False) & "). Проверьте их вручную."
    End If
    MsgBox msg, vbCritical, "Очистка прервана"
End Sub

' Каждое восстановление - отдельно и под Resume Next: одно упавшее не должно оставить остальные.
Private Sub RestoreApp(ByVal su As Boolean, ByVal ee As Boolean, ByVal calc As XlCalculation, ByVal calcChanged As Boolean)
    On Error Resume Next
    If calcChanged Then Application.Calculation = calc
    Application.EnableEvents = ee
    Application.ScreenUpdating = su
End Sub

Private Sub RestoreCell(ByVal c As Range, ByVal oldValue As String, ByVal oldFormat As Variant)
    c.NumberFormat = "@"
    c.Value2 = oldValue
    c.NumberFormat = oldFormat
End Sub

' Текстовые константы из выделения, видимые. SpecialCells на ОДНОЙ ячейке молча расширяется
' на весь лист - поэтому одиночная ячейка проверяется вручную.
Private Function TextConstantsIn(ByVal sel As Range, ByRef hiddenCount As Long) As Range
    Dim r As Range, v As Range
    hiddenCount = 0
    ' Целый столбец или Ctrl+A режутся до использованной области листа
    Set sel = Application.Intersect(sel, sel.Parent.UsedRange)
    If sel Is Nothing Then Exit Function
    If sel.CountLarge = 1 Then
        If sel.HasFormula Then Exit Function
        If VarType(sel.Value2) <> vbString Then Exit Function
        If sel.EntireRow.Hidden Or sel.EntireColumn.Hidden Then hiddenCount = 1: Exit Function
        Set TextConstantsIn = sel
        Exit Function
    End If
    On Error Resume Next
    Set r = sel.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo 0
    If r Is Nothing Then Exit Function
    If r.CountLarge = 1 Then
        If r.EntireRow.Hidden Or r.EntireColumn.Hidden Then hiddenCount = 1: Exit Function
        Set TextConstantsIn = r
        Exit Function
    End If
    On Error Resume Next
    Set v = r.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If v Is Nothing Then hiddenCount = r.CountLarge: Exit Function
    hiddenCount = r.CountLarge - v.CountLarge
    Set TextConstantsIn = v
End Function

' Разное оформление внутри одной ячейки: свойство шрифта возвращает Null, если оно неодинаково
' по символам. Проверяются все свойства, которые можно назначить через Characters(...).Font.
Private Function IsRichText(ByVal c As Range) As Boolean
    With c.Font
        IsRichText = IsNull(.Bold) Or IsNull(.Italic) Or IsNull(.Color) Or IsNull(.Size) _
                     Or IsNull(.Name) Or IsNull(.Underline) Or IsNull(.Strikethrough) _
                     Or IsNull(.Superscript) Or IsNull(.Subscript) Or IsNull(.FontStyle) _
                     Or IsNull(.ThemeColor) Or IsNull(.TintAndShade) Or IsNull(.ThemeFont)
    End With
End Function

' Текст пишется как текст: формат "@" на время записи, потом прежний формат обратно.
' Иначе "00123" станет числом, "12/03" - датой, а "=1+1" - формулой.
Private Sub PutText(ByVal c As Range, ByVal s As String)
    Dim f As Variant
    f = c.NumberFormat
    c.NumberFormat = "@"
    c.Value2 = s
    c.NumberFormat = f
End Sub

Private Function IsCyrillic(ch As String) As Boolean
    Dim code As Long
    code = AscW(ch) And &HFFFF&
    IsCyrillic = (code >= &H400 And code <= &H4FF)
End Function

' Омоглифы меняются только в строке, похожей на код: есть цифра, есть латинская буква,
' латинских букв не меньше кириллических, и вся кириллица - из набора омоглифов.
' Русское слово ("ВЕТЕР", "СМР-2", "МОСТ-A") остаётся как есть и считается как неоднозначное.
Private Function FixHomoglyphs(s As String, ByRef st As Stats) As String
    Dim i As Long, ch As String, nLatin As Long, nCyr As Long, hasDigit As Boolean, n As Long
    FixHomoglyphs = s
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            nLatin = nLatin + 1
        ElseIf ch >= "0" And ch <= "9" Then
            hasDigit = True
        ElseIf IsCyrillic(ch) Then
            nCyr = nCyr + 1
            If InStr(HOMO_FROM, ch) = 0 Then
                st.ambiguous = st.ambiguous + 1
                Exit Function
            End If
        End If
    Next i
    If nCyr = 0 Then Exit Function
    If Not hasDigit Or nLatin = 0 Or nLatin < nCyr Then
        st.ambiguous = st.ambiguous + 1
        Exit Function
    End If
    For i = 1 To Len(HOMO_FROM)
        ch = Mid$(HOMO_FROM, i, 1)
        n = Len(s) - Len(Replace(s, ch, vbNullString))
        If n > 0 Then
            st.homoglyphs = st.homoglyphs + n
            s = Replace(s, ch, Mid$(HOMO_TO, i, 1))
        End If
    Next i
    FixHomoglyphs = s
End Function

Private Function CountOf(s As String, what As String) As Long
    CountOf = (Len(s) - Len(Replace(s, what, vbNullString))) \ Len(what)
End Function

Private Function FixLineBreaks(s As String, ByRef st As Stats) As String
    st.lineBreaks = st.lineBreaks + CountOf(s, vbCrLf)
    s = Replace(s, vbCrLf, " ")
    st.lineBreaks = st.lineBreaks + CountOf(s, vbCr) + CountOf(s, vbLf) + CountOf(s, ChrW(&H2028)) + CountOf(s, ChrW(&H2029))
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ChrW(&H2028), " ")
    s = Replace(s, ChrW(&H2029), " ")
    FixLineBreaks = s
End Function

' Пробельные символы, которые приходят из веба, SAP и PDF и выглядят как обычный пробел:
' табуляция, неразрывный (A0), огамский (1680), U+2000-200A, узкий неразрывный (202F),
' математический (205F), идеографический (3000) - список пробелов Unicode (White_Space). Символы нулевой ширины (200B-200D, 2060, FEFF) не видны вовсе - удаляются.
Private Function OtherSpacesToSpace(s As String) As String
    Dim code As Long
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW(&HA0), " ")
    s = Replace(s, ChrW(&H1680), " ")
    For code = &H2000 To &H200A
        s = Replace(s, ChrW(code), " ")
    Next code
    s = Replace(s, ChrW(&H202F), " ")
    s = Replace(s, ChrW(&H205F), " ")
    s = Replace(s, ChrW(&H3000), " ")
    For code = &H200B To &H200D
        s = Replace(s, ChrW(code), vbNullString)
    Next code
    s = Replace(s, ChrW(&H2060), vbNullString)
    s = Replace(s, ChrW(&HFEFF), vbNullString)
    OtherSpacesToSpace = s
End Function

Private Function RemoveSpaces(s As String, ByRef st As Stats) As String
    Dim before As Long
    before = Len(s)
    s = Replace(OtherSpacesToSpace(s), " ", vbNullString)
    st.spacesRemoved = st.spacesRemoved + (before - Len(s))
    RemoveSpaces = s
End Function

Private Function NormalizeSpaces(s As String, ByRef st As Stats) As String
    Dim before As Long, prevLen As Long
    before = Len(s)
    s = OtherSpacesToSpace(s)
    Do
        prevLen = Len(s)
        s = Replace(s, "  ", " ")
    Loop While Len(s) <> prevLen
    s = Trim$(s)
    st.spacesCollapsed = st.spacesCollapsed + (before - Len(s))
    NormalizeSpaces = s
End Function

Private Function Report(st As Stats, c As Boolean, l As Boolean, r As Boolean, n As Boolean) As String
    Dim m As String
    m = "Обработка завершена." & vbCrLf & vbCrLf
    m = m & "Текстовых ячеек: " & st.cellsSeen & ", изменено: " & st.cellsChanged & vbCrLf
    If st.hidden > 0 Then m = m & "Пропущено скрытых: " & st.hidden & vbCrLf
    If st.richText > 0 Then m = m & "Пропущено с разным оформлением символов: " & st.richText & vbCrLf
    m = m & vbCrLf
    If c Then
        m = m & ChrW(&H2713) & " Кириллица: заменено символов " & st.homoglyphs
        If st.ambiguous > 0 Then m = m & "; не тронуто ячеек с кириллицей (русский текст или код без цифр/латиницы): " & st.ambiguous
        m = m & vbCrLf
    End If
    If l Then m = m & ChrW(&H2713) & " Переносов строк заменено: " & st.lineBreaks & vbCrLf
    If r Then m = m & ChrW(&H2713) & " Пробельных и невидимых символов удалено: " & st.spacesRemoved & vbCrLf
    If n Then m = m & ChrW(&H2713) & " Пробельных и невидимых символов убрано при нормализации: " & st.spacesCollapsed & vbCrLf
    If st.cellsChanged = 0 Then m = m & vbCrLf & "Изменений не найдено."
    Report = m
End Function
