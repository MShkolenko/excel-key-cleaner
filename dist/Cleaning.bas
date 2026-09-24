Attribute VB_Name = "Cleaning"
Option Explicit
Option Compare Binary   ' омоглифы и Replace зависят от регистра - не давать проекту переопределить

' Очистка ключей в ВЫДЕЛЕННЫХ ячейках: кириллические омоглифы -> латиница, переносы строк,
' невидимые символы, пробелы. Правило версии 3: ни одна текстовая ячейка выделения не остаётся
' необработанной молча. Ячейки с разным оформлением символов правятся посимвольно, оформление
' сохраняется. Всё, что макрос НЕ тронул, посчитано в отчёте; на что стоит посмотреть - код, за
' который правило не взялось, объединённая область не целиком, заголовок таблицы, - выделено на
' листе (выделение проверяется, а не предполагается); адресов в отчёте нет. Ошибки не глотаются. Текст
' остаётся текстом - и это проверяется чтением после записи. План считается до первой записи;
' при ошибке (в том числе Ctrl+Break) уже записанное откатывается, rich text - из точной копии.

' У/у намеренно НЕ в списке: русская у и латинская Y - не одно и то же (оператор, 2026-09-21).
Private Const HOMO_FROM As String = "АВЕКМНОРСТХаеорсх"
Private Const HOMO_TO As String = "ABEKMHOPCTXaeopcx"

Private Const MAX_LISTED As Long = 12      ' адресов в аварийном сообщении о сорвавшемся откате; остальные - "и ещё N"
' Измерено (Excel 2024): правки через Range.Characters на тексте длиннее 255 символов - молчаливый
' no-op (255 работает, 256 - нет). Такие rich-ячейки пишутся целиком и перечисляются в отчёте.
Private Const RICH_EDIT_MAX As Long = 255
Private Const REPORT_BUDGET As Long = 1000 ' MsgBox обрезает ~1024 символа; исключения важнее статистики
' Потолки аварийного блока: 200 + 60 + 120 плюс ~300 символов постоянного текста - он всегда влезает
' в окно целиком, каким бы длинным ни было описание ошибки от Excel.
Private Const ALARM_ERR_MAX As Long = 200
Private Const ALARM_NAME_MAX As Long = 60
Private Const ALARM_STATE_MAX As Long = 120

Private Type Stats
    selected As Double
    textCells As Long
    formulas As Long
    numbers As Long
    others As Long
    empties As Double
    cellsChanged As Long
    richKept As Long        ' rich text, оформление сохранено
    richFlattened As Long   ' rich text, записан целиком: Excel сохраняет разметку ПО ПОЗИЦИЯМ, она могла съехать
    crlfToLf As Long        ' CR, которые Excel сам убрал при посимвольной правке (CRLF -> LF, перенос остался)
    hiddenText As Long      ' текстовых ячеек в скрытых строках/столбцах (обрабатываются наравне с видимыми)
    homoglyphs As Long
    lineBreaks As Long
    spacesRemoved As Long
    spacesCollapsed As Long
    edgesTrimmed As Long    ' пробелы и невидимые по краям - убираются ВСЕГДА, независимо от флажков
    cyrLeft As Long         ' ячеек, где кириллица осталась (всего)
    cyrMixed As Long        ' из них подозрительных: латиница и кириллица в одном слове - выделяются
    mergedPartial As Long
    tableHeaders As Long
End Type

Private Const KIND_PLAIN As Long = 0
Private Const KIND_RICH As Long = 1

' Ячейки для итогового выделения, собранные полосами: соседние по вертикали ячейки одного столбца
' склеиваются сразу при добавлении, Union строится один раз в конце (RunsToRange).
Private Type CellRuns
    sh As Worksheet
    col() As Long
    top() As Long
    bottom() As Long
    lastRun() As Long       ' по номеру столбца: индекс последней полосы этого столбца
    n As Long
    capacity As Long
End Type

' Единственная точка входа. Прежнее имя Замена_Кирилицы_С_Выбором убрано по просьбе владельца
' (2026-09-22): макрос показывался в списке Excel дважды. Кнопке, назначенной на старое имя, надо
' один раз переназначить макрос на CleanKeys.
Public Sub CleanKeys()
    Dim doCyrillic As Boolean, doLineBreaks As Boolean, doRemoveSpaces As Boolean, doNormalizeSpaces As Boolean
    Dim onlyCodes As Boolean
    Dim sel As Range, target As Range, cell As Range
    Dim st As Stats
    Dim plannedCells() As Range, plannedOld() As String, plannedNew() As String, plannedFmt() As Variant, plannedKind() As Long
    Dim plannedCount As Long, plannedCapacity As Long
    Dim cyrMixedCells As Range, flattenedCells As Range, mergedPartial As Range, headerCells As Range
    Dim cyrMixedRuns As CellRuns, flattenedRuns As CellRuns, mergedRuns As CellRuns, headerRuns As CellRuns
    Dim backup As Workbook, backupRow As Long, plannedBackupRow() As Long, srcBook As Workbook
    Dim s0 As String, s As String, isRich As Boolean, crDrop As Long
    Dim i As Long, attempted As Long, rolledBack As Long, rollbackFailed As Long, richRestoreFailed As Range
    Dim oldSU As Boolean, oldEE As Boolean, oldCalc As XlCalculation, oldCancel As XlEnableCancelKey
    Dim calcChanged As Boolean, appChanged As Boolean
    Dim oldStatusBar As Variant, statusBarShown As Boolean
    Dim errText As String, msg As String, selectedOk As Boolean, exceptions As Range, backupLeftOpen As Boolean
    Dim postErr As String, backupName As String, appLeft As String, alarm As String
    Dim nothingChosen As Boolean
    Dim anyTables As Boolean, anyMerged As Boolean, mergedFlag As Variant

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
    onlyCodes = CleaningForm.Controls("CheckBox5").Value
    Unload CleaningForm

    ' Ни одного действия - это НЕ отказ: края чистятся всегда, и прогон без единого флажка
    ' делает ровно это. Отказаться от прогона можно кнопкой "Отмена" (владелец, 2026-09-22).
    nothingChosen = Not (doCyrillic Or doLineBreaks Or doRemoveSpaces Or doNormalizeSpaces)

    ' ---------- что в выделении: полный учёт, чтобы ни одна ячейка не "исчезла" ----------
    Set sel = Application.Intersect(Selection, Selection.Parent.UsedRange)
    If sel Is Nothing Then
        MsgBox "В выделении нет заполненных ячеек.", vbInformation
        Exit Sub
    End If
    Set srcBook = sel.Parent.Parent
    st.selected = sel.CountLarge
    st.formulas = CountOfType(sel, xlCellTypeFormulas, 0)
    st.numbers = CountOfType(sel, xlCellTypeConstants, xlNumbers)
    st.others = CountOfType(sel, xlCellTypeConstants, xlLogical + xlErrors)
    On Error GoTo PlanFail
    Set target = TextConstantsIn(sel, NonEmptyCount(sel) - st.formulas - st.numbers - st.others)
    On Error GoTo 0
    If target Is Nothing Then
        MsgBox "В выделении нет текстовых ячеек - чистить нечего." & vbCrLf & _
               WhatWasThere(st), vbInformation, "Результат обработки"
        Exit Sub
    End If
    st.textCells = target.CountLarge
    st.empties = st.selected - st.textCells - st.formulas - st.numbers - st.others
    ' три вопроса ко всему выделению сразу, чтобы не задавать их каждой ячейке
    st.hiddenText = HiddenCount(target)
    anyTables = (sel.Parent.ListObjects.Count > 0)
    mergedFlag = sel.MergeCells                 ' False = нет объединений, True = все, Null = часть
    anyMerged = IsNull(mergedFlag)
    If Not anyMerged Then anyMerged = (mergedFlag = True)

    ' ---------- план: что во что превратится, без единой записи ----------
    oldStatusBar = Application.StatusBar
    statusBarShown = True
    On Error GoTo PlanFail
    i = 0
    For Each cell In target
        i = i + 1
        ShowProgress "проверка", i, CDbl(st.textCells)
        If anyTables And IsTableHeader(cell) Then
            ' Excel сам переименовывает дубликаты заголовков и правит структурные ссылки в формулах
            ' вне выделения - записать сюда значит получить не то, что запланировано, и тронуть чужое
            st.tableHeaders = st.tableHeaders + 1
            Collect headerRuns, cell
            GoTo NextCell
        End If
        ' MergeArea, а не MergeCells: чтение cell.MergeCells по ячейке копит GDI-объекты Excel (замер
        ' 2026-09-24, копия рабочей книги 33 053 x 52: +9 562 за проход, 160 с, до потолка 10 000 -
        ' после чего Workbooks.Add для резервной копии отказывал, и прогон откатывался целиком).
        ' MergeArea.CountLarge на тех же ячейках: GDI не растёт, 1 с, те же 38 объединённых ячеек.
        If anyMerged And cell.MergeArea.CountLarge > 1 Then
            If Application.Intersect(cell.MergeArea, sel).CountLarge <> cell.MergeArea.CountLarge Then
                st.mergedPartial = st.mergedPartial + 1
                Collect mergedRuns, cell
                GoTo NextCell
            End If
        End If
        s0 = cell.Value2
        s = s0
        If doCyrillic Then s = FixHomoglyphs(s, st, onlyCodes, cell, cyrMixedRuns)
        If doLineBreaks Then s = FixLineBreaks(s, st)
        If doRemoveSpaces Then s = RemoveSpaces(s, st)
        If doNormalizeSpaces Then s = NormalizeSpaces(s, st)
        s = TrimEdges(s, st, doLineBreaks)   ' края - ВСЕГДА, ни от какого флажка не зависит
        If s <> s0 Then
            If cell.Parent.ProtectContents And cell.Locked Then
                RestoreStatusBar oldStatusBar, statusBarShown
                MsgBox "Ячейка " & cell.Address(False, False) & " заблокирована на защищённом листе." & vbCrLf & _
                       "Ничего не изменено.", vbCritical, "Защищённый лист"
                Exit Sub
            End If
            isRich = IsRichText(cell)
            If plannedCount = plannedCapacity Then GrowPlanned plannedCells, plannedOld, plannedNew, plannedFmt, plannedKind, plannedCapacity
            plannedCount = plannedCount + 1
            Set plannedCells(plannedCount) = cell
            plannedOld(plannedCount) = s0
            plannedNew(plannedCount) = s
            plannedFmt(plannedCount) = cell.NumberFormat
            plannedKind(plannedCount) = IIf(isRich, KIND_RICH, KIND_PLAIN)
        End If
NextCell:
    Next cell
    On Error GoTo 0

    ' ---------- запись с откатом ----------
    ' Обработчик взводится ДО первого касания Application.*: даже смена режима пересчёта может
    ' упасть (книга с OLAP-источником допускает только автоматический режим). Ctrl+Break тоже
    ' приходит сюда как ошибка (EnableCancelKey = xlErrorHandler), а не бросает книгу полуправленной.
    attempted = 0
    On Error GoTo Fail
    oldSU = Application.ScreenUpdating
    oldEE = Application.EnableEvents
    oldCalc = Application.Calculation
    oldCancel = Application.EnableCancelKey
    appChanged = True
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    calcChanged = True

    ' Rich text откатывается ТОЧНО: до первой записи каждая такая ячейка (вся объединённая область,
    ' если она объединена) копируется во временную скрытую книгу и при ошибке копируется обратно.
    ReDim plannedBackupRow(1 To IIf(plannedCount > 0, plannedCount, 1))
    backupRow = 0
    For i = 1 To plannedCount
        If plannedKind(i) = KIND_RICH Then
            If backup Is Nothing Then
                Set backup = Application.Workbooks.Add(xlWBATWorksheet)
                backup.Windows(1).Visible = False
                srcBook.Activate
            End If
            backupRow = backupRow + 1
            plannedCells(i).MergeArea.Copy Destination:=backup.Worksheets(1).Cells(backupRow, 1)
            plannedBackupRow(i) = backupRow
            backupRow = backupRow + plannedCells(i).MergeArea.Rows.Count   ' следующая копия ниже, с зазором
        End If
    Next i

    For i = 1 To plannedCount
        attempted = i
        ShowProgress "запись", i, CDbl(plannedCount)
        If plannedKind(i) = KIND_RICH Then
            crDrop = 0
            If EditPreservingFormat(plannedCells(i), plannedOld(i), plannedNew(i), crDrop) Then
                st.richKept = st.richKept + 1
                st.crlfToLf = st.crlfToLf + crDrop
            Else
                ' Посимвольно не сошлось - пишем целиком и НАЗЫВАЕМ ячейку. Замерено: целая запись
                ' не сбрасывает разметку символов, Excel оставляет её ПО ПОЗИЦИЯМ - а значит после
                ' сдвига текста она может оказаться на других символах. Поэтому это исключение, но
                ' формулировка "сброшено" была бы неправдой.
                PutText plannedCells(i), plannedNew(i)
                st.richFlattened = st.richFlattened + 1
                Collect flattenedRuns, plannedCells(i)
            End If
        Else
            PutText plannedCells(i), plannedNew(i)
        End If
    Next i
    st.cellsChanged = attempted
    ' Записи завершены: откатывать больше нечего - но состояние Excel и резервную книгу обязаны
    ' прибрать ЛЮБЫМ путём, поэтому управление переходит к своему обработчику ОДНОЙ инструкцией.
    ' Раньше здесь стояло "On Error GoTo 0", и между ним и отключением Ctrl+Break было окно, в
    ' котором не действовал ни один обработчик: прерывание в нём оставляло книгу с выключенными
    ' событиями, ручным пересчётом и открытой скрытой резервной книгой.
    On Error GoTo PostWrite
    Application.EnableCancelKey = xlDisabled

    ' Исключения выделяются, пока события ещё выключены (чужой Worksheet_SelectionChange не может
    ' перехватить), и результат ПРОВЕРЯЕТСЯ - слово "выделены" в отчёте появится только по факту.
    Set cyrMixedCells = RunsToRange(cyrMixedRuns)
    Set mergedPartial = RunsToRange(mergedRuns)
    Set headerCells = RunsToRange(headerRuns)
    Set flattenedCells = RunsToRange(flattenedRuns)
    Set exceptions = UnionOf(cyrMixedCells, mergedPartial, headerCells, flattenedCells)
    selectedOk = SelectVerified(exceptions)
    If Not backup Is Nothing Then backupName = backup.Name
    backupLeftOpen = Not CloseBackup(backup)

PostDone:
    On Error Resume Next
    RestoreStatusBar oldStatusBar, statusBarShown
    appLeft = RestoreApp(oldSU, oldEE, oldCalc, oldCancel, calcChanged, appChanged)
    On Error GoTo 0
    msg = Report(st, doCyrillic, doLineBreaks, doRemoveSpaces, doNormalizeSpaces, onlyCodes, _
                 flattenedCells, mergedPartial, headerCells, selectedOk, _
                 backupLeftOpen, backupName, appLeft, postErr, nothingChosen)
    ' Аварийные предупреждения стоят сразу после заголовка, но если отчёт ВСЁ РАВНО длиннее окна,
    ' MsgBox обрежет хвост - поэтому сначала отдельное окно только с ними. Лишний щелчок здесь
    ' дешевле, чем невидимое "закройте временную книгу вручную".
    alarm = ReportAlarm(backupLeftOpen, backupName, appLeft, postErr)
    If Len(alarm) > 0 And Len(msg) > REPORT_BUDGET Then _
        MsgBox "Отчёт целиком в окно не помещается, поэтому главное - отдельно:" & vbCrLf & alarm, _
               vbExclamation, "Требуется внимание"
    MsgBox msg, IIf(Len(alarm) > 0, vbExclamation, vbInformation), "Результат обработки"
    Exit Sub

PostWrite:
    ' Ошибка ПОСЛЕ последней записи: ячейки уже очищены и перечитаны, откатывать нечего и незачем -
    ' но прибраться за собой обязаны, поэтому сюда, а не в Fail. Отчёт скажет об этом вслух.
    postErr = Err.Description
    If Err.Number = 18 Then postErr = "прервано пользователем (Ctrl+Break) уже после записи"
    On Error Resume Next
    Application.EnableCancelKey = xlDisabled
    If Not backup Is Nothing Then
        backupName = backup.Name
        backupLeftOpen = Not CloseBackup(backup)
    End If
    Resume PostDone

PlanFail:
    errText = Err.Description
    RestoreStatusBar oldStatusBar, statusBarShown
    MsgBox "Ошибка при подготовке: " & errText & vbCrLf & vbCrLf & "Ничего не изменено.", vbCritical, "Очистка прервана"
    Exit Sub

Fail:
    errText = Err.Description
    If Err.Number = 18 Then errText = "прервано пользователем (Ctrl+Break)"
    ' Откат покрывает и ячейку, на которой упала запись. Каждая возвращённая ячейка ПЕРЕЧИТЫВАЕТСЯ:
    ' "возвращено" здесь значит "прочитано обратно и совпало", а не "запись не бросила ошибку".
    On Error Resume Next
    Application.EnableCancelKey = xlDisabled   ' откат нельзя прервать посередине
    For i = attempted To 1 Step -1
        If RestoreOne(plannedCells(i), plannedKind(i), plannedOld(i), plannedFmt(i), backup, plannedBackupRow(i)) Then
            rolledBack = rolledBack + 1
        Else
            rollbackFailed = rollbackFailed + 1
            If plannedKind(i) = KIND_RICH Then AddTo richRestoreFailed, plannedCells(i)
        End If
    Next i
    If rollbackFailed = 0 Then
        If Not CloseBackup(backup) Then msg = "Временную книгу " & backup.Name & " закрыть не удалось - закройте вручную, не сохраняя." & vbCrLf
    ElseIf Not backup Is Nothing Then
        ' резервная книга нужна пользователю - показать, не закрывать
        backup.Windows(1).Visible = True
        srcBook.Activate
    End If
    RestoreStatusBar oldStatusBar, statusBarShown
    appLeft = RestoreApp(oldSU, oldEE, oldCalc, oldCancel, calcChanged, appChanged)
    msg = msg & "Ошибка при записи: " & errText & vbCrLf & vbCrLf
    If attempted = 0 Then
        msg = msg & "Ни одна ячейка не изменена."
    ElseIf rollbackFailed = 0 Then
        msg = msg & "Затронутые ячейки (" & attempted & ") возвращены к исходным значениям - каждая прочитана обратно и совпала."
    Else
        msg = msg & "Возвращено и проверено: " & rolledBack & ". НЕ удалось вернуть: " & rollbackFailed & _
              " (первая из затронутых: " & plannedCells(1).Address(False, False) & ", последняя: " & _
              plannedCells(attempted).Address(False, False) & "). Проверьте их вручную."
        If Not richRestoreFailed Is Nothing And Not backup Is Nothing Then
            msg = msg & vbCrLf & "Исходные копии ячеек с оформлением оставлены в книге " & backup.Name & _
                  " (лист 1, столбец A): " & ListAddresses(richRestoreFailed, MAX_LISTED)
        End If
    End If
    If Len(appLeft) > 0 Then msg = msg & vbCrLf & vbCrLf & AppLeftText(appLeft)
    MsgBox msg, vbCritical, "Очистка прервана"
End Sub

' ---------------------------------------------------------------------------------------------
' Отбор и учёт

' Текстовые константы выделения - ВСЕ, включая скрытые строки. SpecialCells на ОДНОЙ ячейке
' молча расширяется на весь лист - поэтому одиночная ячейка проверяется вручную.
' SpecialCells сообщает "ячеек не найдено" той же ошибкой 1004, что и любую другую беду, поэтому
' пустому результату верим только если независимый счёт тоже даёт ноль текстовых ячеек.
Private Function TextConstantsIn(ByVal sel As Range, ByVal expectedText As Double) As Range
    Dim r As Range, errNum As Long, errDesc As String
    If sel.CountLarge = 1 Then
        If sel.HasFormula Then Exit Function
        If VarType(sel.Value2) <> vbString Then Exit Function
        Set TextConstantsIn = sel
        Exit Function
    End If
    On Error Resume Next
    Set r = sel.SpecialCells(xlCellTypeConstants, xlTextValues)
    errNum = Err.Number: errDesc = Err.Description
    On Error GoTo 0
    If r Is Nothing Then
        If expectedText > 0 Then
            Err.Raise IIf(errNum <> 0, errNum, vbObjectError + 1), "TextConstantsIn", _
                      "SpecialCells не вернул текстовые ячейки, хотя по счёту их " & Format$(expectedText, "#,##0") & _
                      IIf(Len(errDesc) > 0, ". " & errDesc, "")
        End If
        Exit Function
    End If
    Set TextConstantsIn = r
End Function

' Число непустых ячеек выделения, посчитанное Excel независимо от SpecialCells.
Private Function NonEmptyCount(ByVal sel As Range) As Double
    Dim area As Range
    For Each area In sel.Areas
        NonEmptyCount = NonEmptyCount + Application.WorksheetFunction.CountA(area)
    Next area
End Function

' Сколько текстовых ячеек сидит в скрытых строках/столбцах - одним запросом, не по ячейке.
Private Function HiddenCount(ByVal r As Range) As Long
    Dim v As Range
    If r.CountLarge = 1 Then
        If r.EntireRow.Hidden Or r.EntireColumn.Hidden Then HiddenCount = 1
        Exit Function
    End If
    On Error Resume Next
    Set v = r.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If v Is Nothing Then HiddenCount = r.CountLarge Else HiddenCount = r.CountLarge - v.CountLarge
End Function

' Сколько ячеек данного вида в выделении. Ошибка "не найдено" = 0; на одной ячейке SpecialCells
' расширяется на лист, поэтому одиночная ячейка считается напрямую.
Private Function CountOfType(ByVal sel As Range, ByVal cellType As XlCellType, ByVal valueType As Long) As Long
    Dim r As Range
    If sel.CountLarge = 1 Then
        If cellType = xlCellTypeFormulas Then
            If sel.HasFormula Then CountOfType = 1
        ElseIf Not sel.HasFormula Then
            Select Case VarType(sel.Value2)
                Case vbDouble, vbCurrency, vbDate, vbInteger, vbLong, vbSingle
                    If valueType = xlNumbers Then CountOfType = 1
                Case vbBoolean, vbError
                    If valueType = xlLogical + xlErrors Then CountOfType = 1
            End Select
        End If
        Exit Function
    End If
    On Error Resume Next
    If valueType = 0 Then
        Set r = sel.SpecialCells(cellType)
    Else
        Set r = sel.SpecialCells(cellType, valueType)
    End If
    On Error GoTo 0
    If Not r Is Nothing Then CountOfType = r.CountLarge
End Function

Private Function IsTableHeader(ByVal c As Range) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = c.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If lo.HeaderRowRange Is Nothing Then Exit Function
    IsTableHeader = Not Application.Intersect(c, lo.HeaderRowRange) Is Nothing
End Function

' Разное оформление внутри одной ячейки: свойство шрифта возвращает Null, если оно неодинаково
' по символам. ThemeColor/ThemeFont намеренно НЕ проверяются: они бросают ошибку на ЛЮБОЙ ячейке
' с вручную заданным цветом, даже однородным (v2.3 -> v2.4). Если какое-то свойство всё же
' упало - ячейка идёт по посимвольному пути, который оформление в любом случае сохраняет.
Private Function IsRichText(ByVal c As Range) As Boolean
    On Error GoTo Unclear
    With c.Font
        IsRichText = IsNull(.Bold) Or IsNull(.Italic) Or IsNull(.Color) Or IsNull(.Size) _
                     Or IsNull(.Name) Or IsNull(.Underline) Or IsNull(.Strikethrough) _
                     Or IsNull(.Superscript) Or IsNull(.Subscript) Or IsNull(.FontStyle) _
                     Or IsNull(.TintAndShade)
    End With
    Exit Function
Unclear:
    IsRichText = True
End Function

' ---------------------------------------------------------------------------------------------
' Запись

' Текст пишется как текст: формат "@" на время записи, потом прежний формат обратно. Иначе "00123"
' станет числом, "12/03" - датой, а "=1+1" - формулой. Записанное ПЕРЕЧИТЫВАЕТСЯ: если Excel положил
' в ячейку не то, что просили, это ошибка, а не тихий успех.
Private Sub PutText(ByVal c As Range, ByVal s As String)
    Dim f As Variant
    f = c.NumberFormat
    c.NumberFormat = "@"
    c.Value2 = s
    c.NumberFormat = f
    Dim back As Variant
    back = c.Value2
    If Len(s) = 0 Then
        ' ячейка из одних пробелов/невидимых символов очищается до пустой - это и есть ожидаемый итог
        If Not (IsEmpty(back) Or (VarType(back) = vbString And Len(back) = 0)) Then _
            Err.Raise vbObjectError + 3, "PutText", "ячейка " & c.Address(False, False) & " после очистки не пуста"
        Exit Sub
    End If
    If VarType(back) <> vbString Then Err.Raise vbObjectError + 2, "PutText", "ячейка " & c.Address(False, False) & " после записи перестала быть текстом"
    If back <> s Then Err.Raise vbObjectError + 3, "PutText", "ячейка " & c.Address(False, False) & " после записи содержит не то, что записано"
End Sub

' Возврат одной ячейки при откате. True только если после возврата ячейка прочитана и совпала.
Private Function RestoreOne(ByVal c As Range, ByVal kind As Long, ByVal oldValue As String, ByVal oldFormat As Variant, _
                            ByVal backup As Workbook, ByVal backupRow As Long) As Boolean
    On Error Resume Next
    If kind = KIND_RICH And Not backup Is Nothing And backupRow > 0 Then
        Err.Clear
        backup.Worksheets(1).Cells(backupRow, 1).MergeArea.Copy Destination:=c.MergeArea.Cells(1, 1)
        If Err.Number = 0 Then
            If c.Value2 = oldValue Then RestoreOne = True: Exit Function
        End If
    End If
    Err.Clear
    c.NumberFormat = "@"
    c.Value2 = oldValue
    c.NumberFormat = oldFormat
    If Err.Number <> 0 Then Exit Function
    RestoreOne = (c.Value2 = oldValue) And (kind = KIND_PLAIN)   ' rich, вернувшийся целой записью, - не точный возврат
End Function

' Посимвольная правка с сохранением оформления. Все преобразования макроса - замена одного символа
' на один (омоглиф, перенос/экзотический пробел -> пробел) или удаление (нулевая ширина, убранные
' пробелы, схлопнутые повторы, обрезанные края). Вставок не бывает.
' Обход по сериям: серия пробельных символов старого текста ставится в соответствие серии пробелов
' нового; внутри серии сначала СОХРАНЯЮТСЯ исходные пробелы (с их оформлением), потом подставляются
' символы, которым макрос дал пробел, остальное удаляется. Так оформленный пробел не проигрывает
' бывшему переносу строки (замечание Codex по X<CR><LF><пробел>Y). Правки применяются справа
' налево через Range.Characters. Возвращает False, если выравнивание не сошлось или Characters
' бросил ошибку - тогда вызывающий пишет ячейку целиком и сообщает об этом.
' Измерено: Characters нумерует символы по тексту, где CRLF - ОДИН символ (Value2 отдаёт два),
' поэтому выравнивание идёт по тексту с CRLF -> LF; после правок ячейка перечитывается.
' Правит rich-ячейку посимвольно (Characters), сохраняя оформление каждого символа. False = старый
' текст нельзя объяснить новым (ячейка будет записана целиком) или контрольное чтение не совпало.
' Измерено: Characters нумерует CRLF как ОДИН символ и после любой правки оставляет только LF -
' поэтому обе стороны выравниваются по тексту с CRLF -> LF, а число убранных Excel'ем CR
' возвращается в crDropped, чтобы отчёт сказал об этом (перенос строки при этом остаётся).
Private Function EditPreservingFormat(ByVal c As Range, ByVal oldText As String, ByVal newText As String, ByRef crDropped As Long) As Boolean
    Dim ops() As Long, subs() As String   ' ops(i): 0 = оставить, 1 = заменить на subs(i), 2 = удалить
    Dim i As Long, j As Long, n As Long, m As Long, ch As String, want As String, aligned As String, got As String
    Dim runStart As Long, runEnd As Long, k As Long, need As Long, t As Long
    If Len(oldText) > RICH_EDIT_MAX Then Exit Function
    oldText = Replace(oldText, vbCrLf, vbLf)   ' так эту ячейку нумерует Characters
    aligned = Replace(newText, vbCrLf, vbLf)   ' и такой она будет прочитана после правки
    n = Len(oldText): m = Len(aligned)
    If n = 0 Then Exit Function
    ReDim ops(1 To n): ReDim subs(1 To n)
    j = 1
    i = 1
    Do While i <= n
        ch = Mid$(oldText, i, 1)
        If IsSpaceLike(ch) Then
            ' серия пробельных: [runStart, runEnd]
            runStart = i: runEnd = i
            Do While runEnd < n
                If Not IsSpaceLike(Mid$(oldText, runEnd + 1, 1)) Then Exit Do
                runEnd = runEnd + 1
            Loop
            ' серия пробельных символов нового текста с позиции j (пробелы, а также табуляция/перенос,
            ' которые остались на месте, если их действие выключено)
            need = 0
            Do While j + need <= m
                If Not IsSpaceLike(Mid$(aligned, j + need, 1)) Then Exit Do
                need = need + 1
            Loop
            ' обход двух серий по порядку: оставить равный символ; иначе, если ДАЛЬШЕ в старой серии есть
            ' равный нужному и старых символов ещё хватает, удалить этот (исходный символ уносит с собой
            ' своё оформление, а замена приписала бы оформление соседу); иначе заменить, если макрос мог
            ' его так превратить; иначе не сошлось
            k = runStart: t = 0
            Do While t < need
                If k > runEnd Then Exit Function
                want = Mid$(aligned, j + t, 1)
                ch = Mid$(oldText, k, 1)
                If ch = want Then
                    ops(k) = 0: k = k + 1: t = t + 1
                ElseIf EqualAhead(oldText, k + 1, runEnd, want, need - t) Then
                    ops(k) = 2: k = k + 1
                ElseIf ExpectedReplacement(ch) = want Then
                    ops(k) = 1: subs(k) = want: k = k + 1: t = t + 1
                Else
                    ops(k) = 2: k = k + 1
                End If
            Loop
            Do While k <= runEnd                                         ' хвост серии - удалить
                ops(k) = 2: k = k + 1
            Loop
            j = j + need
            i = runEnd + 1
        Else
            If j <= m Then
                If ch = Mid$(aligned, j, 1) Then
                    ops(i) = 0: j = j + 1
                Else
                    want = ExpectedReplacement(ch)
                    If Len(want) = 1 And want = Mid$(aligned, j, 1) Then
                        ops(i) = 1: subs(i) = want: j = j + 1
                    Else
                        Exit Function                                    ' буква, которую нельзя ни оставить, ни заменить
                    End If
                End If
            Else
                Exit Function                                            ' старый текст длиннее, чем объяснимо
            End If
            i = i + 1
        End If
    Loop
    If j <> m + 1 Then Exit Function
    On Error GoTo Bad
    For i = n To 1 Step -1
        Select Case ops(i)
            Case 1: c.Characters(i, 1).Text = subs(i)
            Case 2: c.Characters(i, 1).Delete
        End Select
    Next i
    On Error GoTo 0
    got = c.Value2
    If Replace(got, vbCrLf, vbLf) <> aligned Then Exit Function   ' контроль по факту, а не по расчёту
    crDropped = crDropped + (CountOf(newText, vbCr) - CountOf(got, vbCr))
    EditPreservingFormat = True
    Exit Function
Bad:
    If Err.Number = 18 Then Err.Raise 18   ' Ctrl+Break - это не "не сошлось", это откат всего
    EditPreservingFormat = False
End Function

' Есть ли в старой серии на [fromPos, toPos] символ, равный нужному, причём так, что от него до конца
' серии старых символов хватает на оставшиеся нужные. Первый равный - самый выгодный, дальше только хуже.
Private Function EqualAhead(ByRef s As String, ByVal fromPos As Long, ByVal toPos As Long, ByVal want As String, ByVal remaining As Long) As Boolean
    Dim q As Long
    For q = fromPos To toPos
        If Mid$(s, q, 1) = want Then
            EqualAhead = (toPos - q + 1) >= remaining
            Exit Function
        End If
    Next q
End Function

' Символы, которые действия 2-4 считают пробельными: обычный пробел, всё, что превращается в пробел,
' и всё, что удаляется как невидимое.
Private Function IsSpaceLike(ByVal ch As String) As Boolean
    Dim code As Long
    If ch = " " Then IsSpaceLike = True: Exit Function
    code = AscW(ch) And &HFFFF&
    Select Case code
        ' &HFEFF& с суффиксом Long НЕ случайно: четырёхзначный hex-литерал в VBA - это Integer
        ' со знаком, и &HFEFF равен -257, а не 65279. Без суффикса BOM никогда не попадал в этот
        ' список, и стенд поймал это только на обрезке краёв (2026-09-22).
        Case 9, 10, 11, 12, 13, &H85, &HA0, &H1680, &H2000 To &H200D, &H2028, &H2029, &H202F, &H205F, &H2060, &H3000, &HFEFF&
            IsSpaceLike = True
    End Select
End Function

' Во что макрос может превратить один символ (замена 1:1). Пусто = только удаление.
Private Function ExpectedReplacement(ByVal ch As String) As String
    Dim p As Long, code As Long
    p = InStr(HOMO_FROM, ch)
    If p > 0 Then ExpectedReplacement = Mid$(HOMO_TO, p, 1): Exit Function
    code = AscW(ch) And &HFFFF&
    Select Case code
        Case 9, 10, 11, 12, 13, &H85, &HA0, &H1680, &H2000 To &H200A, &H2028, &H2029, &H202F, &H205F, &H3000
            ExpectedReplacement = " "
    End Select
End Function

' Каждое свойство - отдельно (одно упавшее не должно оставить остальные) и ПЕРЕЧИТЫВАЕТСЯ:
' "восстановлено" здесь значит "прочитано обратно и совпало". Возвращает то, что вернуть НЕ
' удалось (пустая строка = всё на месте) - молча это проглатывать нельзя, пользователь останется
' с выключенными событиями и ручным пересчётом, не зная об этом.
Private Function RestoreApp(ByVal su As Boolean, ByVal ee As Boolean, ByVal calc As XlCalculation, ByVal cancel As XlEnableCancelKey, _
                            ByVal calcChanged As Boolean, ByVal appChanged As Boolean) As String
    Dim bad As String, got As Variant
    If Not appChanged Then Exit Function
    On Error Resume Next
    If calcChanged Then
        Err.Clear: Application.Calculation = calc: got = Application.Calculation
        bad = bad & AppMismatch("режим пересчёта", calc, got)
    End If
    Err.Clear: Application.EnableEvents = ee: got = Application.EnableEvents
    bad = bad & AppMismatch("события", ee, got)
    Err.Clear: Application.ScreenUpdating = su: got = Application.ScreenUpdating
    bad = bad & AppMismatch("обновление экрана", su, got)
    Err.Clear: Application.EnableCancelKey = cancel: got = Application.EnableCancelKey
    bad = bad & AppMismatch("реакция на Ctrl+Break", cancel, got)
    If Len(bad) > 0 Then RestoreApp = Left$(bad, Len(bad) - 2)
End Function

' Имя свойства, если запись упала ИЛИ чтение не подтвердило значение; иначе пусто.
Private Function AppMismatch(ByVal propName As String, ByVal wanted As Variant, ByVal got As Variant) As String
    If Err.Number <> 0 Then AppMismatch = propName & ", ": Exit Function
    If got <> wanted Then AppMismatch = propName & ", "
End Function

Private Function AppLeftText(ByVal bad As String) As String
    AppLeftText = "СОСТОЯНИЕ EXCEL восстановлено не полностью: " & bad & "." & vbCrLf & _
                  "(проверьте Формулы -> Параметры вычислений; если книга ведёт себя странно - перезапустите Excel)"
End Function

' Ход выполнения в строке состояния - без DoEvents (он открыл бы повторный вход в макрос).
Private Sub ShowProgress(ByVal phase As String, ByVal doneCount As Long, ByVal totalCount As Double)
    If doneCount = 1 Or totalCount = 0 Or doneCount Mod 500 = 0 Or CDbl(doneCount) = totalCount Then
        Application.StatusBar = "Очистка ключей - " & phase & ": " & Format$(doneCount, "#,##0") & " / " & Format$(totalCount, "#,##0")
    End If
End Sub

Private Sub RestoreStatusBar(ByVal oldValue As Variant, ByRef shown As Boolean)
    If Not shown Then Exit Sub
    On Error Resume Next
    Application.StatusBar = oldValue
    shown = False
End Sub

' Ёмкость удваивается: суммарная стоимость ReDim Preserve остаётся O(n) амортизированно
' (Collection.Item(i) в VBA - проход по связному списку; 40 000 ячеек: 39 с против 3,9 с, чистый замер).
Private Sub GrowPlanned(ByRef cells() As Range, ByRef olds() As String, ByRef news() As String, ByRef fmts() As Variant, ByRef kinds() As Long, ByRef capacity As Long)
    If capacity = 0 Then
        capacity = 1024
        ReDim cells(1 To capacity): ReDim olds(1 To capacity): ReDim news(1 To capacity)
        ReDim fmts(1 To capacity): ReDim kinds(1 To capacity)
    Else
        capacity = capacity * 2
        ReDim Preserve cells(1 To capacity): ReDim Preserve olds(1 To capacity): ReDim Preserve news(1 To capacity)
        ReDim Preserve fmts(1 To capacity): ReDim Preserve kinds(1 To capacity)
    End If
End Sub

Private Sub AddTo(ByRef acc As Range, ByVal c As Range)
    If acc Is Nothing Then Set acc = c Else Set acc = Application.Union(acc, c)
End Sub

' Union по одной ячейке растёт почти как КУБ от числа областей (замер на Excel 2024, 2026-09-24:
' 1 000 ячеек - 0,3 с, 8 000 - 130 с), а на реальном листе с русским текстом таких ячеек десятки
' тысяч: рабочий лист, 63 109 ячеек - 15 минут. Поэтому ячейки копятся полосами, а Union - один раз, деревом.
Private Sub Collect(ByRef acc As CellRuns, ByVal c As Range)
    Dim k As Long, r As Long, cl As Long
    r = c.Row: cl = c.Column
    If acc.n = 0 Then
        Set acc.sh = c.Parent
        acc.capacity = 256
        ReDim acc.col(1 To acc.capacity): ReDim acc.top(1 To acc.capacity): ReDim acc.bottom(1 To acc.capacity)
        ReDim acc.lastRun(1 To c.Parent.Columns.Count)
    End If
    k = acc.lastRun(cl)
    If k > 0 Then
        If acc.bottom(k) = r - 1 Then acc.bottom(k) = r: Exit Sub    ' продолжение полосы столбца
    End If
    If acc.n = acc.capacity Then
        acc.capacity = acc.capacity * 2
        ReDim Preserve acc.col(1 To acc.capacity): ReDim Preserve acc.top(1 To acc.capacity)
        ReDim Preserve acc.bottom(1 To acc.capacity)
    End If
    acc.n = acc.n + 1
    acc.col(acc.n) = cl: acc.top(acc.n) = r: acc.bottom(acc.n) = r
    acc.lastRun(cl) = acc.n
End Sub

' Полосы -> один Range. Union попарно, уровнями (64 000 областей - 10 с), а не по одной (часы).
Private Function RunsToRange(ByRef acc As CellRuns) As Range
    Dim a() As Range, i As Long, n As Long
    If acc.n = 0 Then Exit Function
    ReDim a(1 To acc.n)
    For i = 1 To acc.n
        Set a(i) = acc.sh.Range(acc.sh.Cells(acc.top(i), acc.col(i)), acc.sh.Cells(acc.bottom(i), acc.col(i)))
    Next i
    n = acc.n
    Do While n > 1
        For i = 1 To n \ 2
            Set a(i) = Application.Union(a(2 * i - 1), a(2 * i))
        Next i
        If n Mod 2 = 1 Then
            Set a(n \ 2 + 1) = a(n)
            n = n \ 2 + 1
        Else
            n = n \ 2
        End If
    Loop
    Set RunsToRange = a(1)
End Function

Private Function UnionOf(ParamArray parts() As Variant) As Range
    Dim i As Long, acc As Range
    For i = LBound(parts) To UBound(parts)
        If Not parts(i) Is Nothing Then
            If acc Is Nothing Then Set acc = parts(i) Else Set acc = Application.Union(acc, parts(i))
        End If
    Next i
    Set UnionOf = acc
End Function

' Выделяет диапазон и ПРОВЕРЯЕТ, что каждая его ячейка действительно выделена. False, если лист
' скрыт, книга не активируется или выделение не то - тогда отчёт не скажет "выделены".
' Строки адреса целиком не сравниваются: при выделении Excel сам сливает соседние области в
' прямоугольники (на рабочем листе: области B2:K2, A1:A7, ... -> выделение B2:K2,A1:K7,...), и
' адрес выделения законно отличается при тех же ячейках - это давало ложное "выделить НЕ удалось".
' Поэтому по областям: каждая область диапазона должна лежать целиком в одной области выделения.
Private Function SelectVerified(ByVal r As Range) As Boolean
    Dim sel As Range, a As Range
    If r Is Nothing Then Exit Function
    On Error GoTo Nope
    r.Parent.Parent.Activate
    r.Parent.Activate
    r.Select
    If TypeName(Selection) <> "Range" Then Exit Function
    Set sel = Selection
    If Not sel.Parent Is r.Parent Then Exit Function
    If sel.Areas.Count = r.Areas.Count Then
        If sel.Address(False, False) = r.Address(False, False) Then SelectVerified = True: Exit Function
    End If
    For Each a In r.Areas
        If Not InsideOneArea(a, sel) Then Exit Function
    Next a
    SelectVerified = True
    Exit Function
Nope:
    SelectVerified = False
End Function

' True, если прямоугольник a целиком лежит в какой-нибудь ОДНОЙ области outer.
Private Function InsideOneArea(ByVal a As Range, ByVal outer As Range) As Boolean
    Dim x As Range, p As Range
    Set x = Application.Intersect(a, outer)
    If x Is Nothing Then Exit Function
    For Each p In x.Areas
        If p.CountLarge = a.CountLarge Then InsideOneArea = True: Exit Function   ' часть a размером с a - это a
    Next p
End Function

' True, если резервной книги нет или она закрылась; False - осталась открытой (ссылка сохраняется).
Private Function CloseBackup(ByRef wb As Workbook) As Boolean
    If wb Is Nothing Then CloseBackup = True: Exit Function
    On Error Resume Next
    wb.Close SaveChanges:=False
    If Err.Number = 0 Then Set wb = Nothing: CloseBackup = True
End Function

' ---------------------------------------------------------------------------------------------
' Преобразования

Private Function IsCyrillic(ch As String) As Boolean
    Dim code As Long
    code = AscW(ch) And &HFFFF&
    IsCyrillic = (code >= &H400 And code <= &H4FF)
End Function

' Омоглифы (v3.1). Решение ПО СЛОВАМ - идея и первая версия: PR #2 (DarthKrya); окончательное
' правило выбрано на копиях пяти рабочих книг (около 1,6 млн ячеек с кириллицей), а не на
' придуманных примерах.
' Слово - кусок между пробелами, разрывами строк и знаками ( ) [ ] { } , ; " « » ' (IsWordBreak).
' Невидимые (U+200B-200D, 2060, FEFF) границей НЕ считаются: макрос их удаляет, а не делает
' пробелом, так что "10U<ZWSP>МА" - одно слово. Дефис, точка и косая черта тоже не границы: на них
' держатся коды (XX.0120.10UМА) и координаты (8Е+15.45/12N+80.45).
' Слово с кириллицей чинится, если правило кода (есть цифра, есть латиница, латинских не меньше
' кириллических, вся кириллица из карты) выполняет:
'  - само слово - код посреди русской фразы: "Опора XX.0120.10UМА" -> "Опора XX.0120.10UMA";
'  - или вся ячейка, как в v3.0 - английский текст, набранный с кириллическими буквами
'    ("Air сompressed 12 bar"), оси "А-В" и коды "(КАА)" в латинской ячейке.
' Русское слово - только кириллица, ни латиницы, ни цифр, и есть СТРОЧНАЯ кириллическая буква - без
' флажка 5 не трогается НИКОГДА: в v3.0 "Насос XX.0120.10UMA" молча становился "Hacoc ...".
' Заглавные слова из одной кириллицы (оси, коды систем) чинятся по правилу ячейки, как в v3.0: на
' рабочих книгах это были только они. Цена: "ВЕТЕР XX.0120.10UMA" (русское слово ЗАГЛАВНЫМИ рядом
' с длинным кодом) по-прежнему станет "BETEP" - как и в v3.0; слово, приклеенное к коду без
' пробела ("Насос-XX.0120.10UMA", "4-х/Safety"), - одно слово, и решение по нему общее.
' Где кириллица осталась, ячейка СЧИТАЕТСЯ (st.cyrLeft). Выделяется только подозрительная (v3.3,
' владелец 2026-09-24: на листе русского текста список из 100 тысяч адресов - шум): латиница и
' кириллица в одном куске слова (куски делит "/", чтобы "кг/kg" не считалось) - это похоже на код,
' за который правило не взялось ("СЕ.U1", "07UУQ", "Cекция"); с флажком 5 - буква без латинской
' пары. Такие идут в st.cyrMixed и в suspRuns. С флажком 5 проверок нет: меняется всё, что есть в карте.
' Замена 1:1, длина строки не меняется - правка на месте через Mid$, строка по символу не склеивается.
Private Function FixHomoglyphs(s As String, ByRef st As Stats, ByVal onlyCodes As Boolean, ByVal c As Range, ByRef suspRuns As CellRuns) As String
    Dim res As String, n As Long, i As Long, j As Long, p As Long, wStart As Long, code As Long
    Dim wLat As Long, wCyr As Long, wDigit As Boolean, wMapped As Boolean, wLower As Boolean, ch As String, w As String
    Dim cLat As Long, cCyr As Long, cDigit As Boolean, cMapped As Boolean, cellOk As Boolean
    Dim leftHere As Boolean, doFix As Boolean, pLat As Boolean, pCyr As Boolean, wMixed As Boolean, suspHere As Boolean
    FixHomoglyphs = s
    n = Len(s)
    ' проход 1 - правило кода по всей ячейке (как в v3.0)
    cMapped = True
    For i = 1 To n
        ch = Mid$(s, i, 1)
        code = AscW(ch) And &HFFFF&
        Select Case code
            Case 65 To 90, 97 To 122: cLat = cLat + 1
            Case 48 To 57: cDigit = True
            Case &H400 To &H4FF
                cCyr = cCyr + 1
                If InStr(HOMO_FROM, ch) = 0 Then cMapped = False
        End Select
    Next i
    If cCyr = 0 Then Exit Function
    cellOk = cDigit And cLat > 0 And cLat >= cCyr And cMapped
    ' проход 2 - решение и правка по каждому слову
    res = s
    wStart = 1: wMapped = True
    For i = 1 To n + 1
        If i <= n Then
            ch = Mid$(s, i, 1)
            code = AscW(ch) And &HFFFF&
            Select Case code
                Case 65 To 90, 97 To 122
                    wLat = wLat + 1
                    pLat = True
                    GoTo NextChar
                Case 48 To 57
                    wDigit = True
                    GoTo NextChar
                Case &H400 To &H4FF
                    wCyr = wCyr + 1
                    pCyr = True
                    If code >= &H430 And code <= &H45F Then wLower = True
                    If InStr(HOMO_FROM, ch) = 0 Then wMapped = False
                    GoTo NextChar
                Case 47                                            ' "/" - граница куска, не слова
                    If pLat And pCyr Then wMixed = True
                    pLat = False: pCyr = False
                    GoTo NextChar
            End Select
            If Not IsWordBreak(ch) Then GoTo NextChar
        End If
        ' конец слова s[wStart .. i-1] (пустое слово между двумя границами ничего не меняет)
        If pLat And pCyr Then wMixed = True
        If wCyr > 0 Then
            If onlyCodes Then
                doFix = True
            ElseIf wLat = 0 And Not wDigit And wLower Then
                doFix = False                                      ' русское слово
            Else
                doFix = cellOk Or (wDigit And wLat > 0 And wLat >= wCyr And wMapped)
            End If
            If doFix Then
                w = Mid$(s, wStart, i - wStart)
                For j = 1 To Len(HOMO_FROM)
                    p = CountOf(w, Mid$(HOMO_FROM, j, 1))
                    If p > 0 Then
                        st.homoglyphs = st.homoglyphs + p
                        w = Replace(w, Mid$(HOMO_FROM, j, 1), Mid$(HOMO_TO, j, 1))
                    End If
                Next j
                Mid$(res, wStart, Len(w)) = w                      ' длина та же: замена 1:1
                If Not wMapped Then                                ' буквы вне карты (Ж, Ш, У...) остались
                    leftHere = True
                    suspHere = True                                ' в слове, которое меняли, - подозрительно
                End If
            Else
                leftHere = True
                If wMixed Then suspHere = True
            End If
        End If
        wStart = i + 1: wLat = 0: wCyr = 0: wDigit = False: wMapped = True: wLower = False
        pLat = False: pCyr = False: wMixed = False
NextChar:
    Next i
    ' ячейка считается ОДИН раз, сколько бы слов в ней ни осталось; выделяется - только подозрительная
    If leftHere Then st.cyrLeft = st.cyrLeft + 1
    If suspHere Then
        st.cyrMixed = st.cyrMixed + 1
        Collect suspRuns, c
    End If
    FixHomoglyphs = res
End Function

' Граница слова для омоглифов: всё, что IsSpaceLike, КРОМЕ невидимых (их макрос удаляет, а не
' превращает в пробел - код с таким мусором внутри остаётся одним словом), и скобки, запятая,
' точка с запятой, кавычки: на рабочих книгах "(ПУ АС)(99UXX)" без них склеивалось в одно слово,
' и русское "АС" становилось латинским.
Private Function IsWordBreak(ByVal ch As String) As Boolean
    Select Case AscW(ch) And &HFFFF&
        Case &H200B To &H200D, &H2060, &HFEFF&
            IsWordBreak = False
        Case 40, 41, 91, 93, 123, 125, 44, 59, 34, 39, &HAB, &HBB   ' ( ) [ ] { } , ; " ' « »
            IsWordBreak = True
        Case Else
            IsWordBreak = IsSpaceLike(ch)
    End Select
End Function

Private Function CountOf(s As String, what As String) As Long
    CountOf = (Len(s) - Len(Replace(s, what, vbNullString))) \ Len(what)
End Function

' Переносы строк: CRLF, CR, LF, вертикальная табуляция (11), разрыв страницы (12), NEL (U+0085),
' U+2028, U+2029 -> пробел.
Private Function FixLineBreaks(s As String, ByRef st As Stats) As String
    Dim code As Variant
    st.lineBreaks = st.lineBreaks + CountOf(s, vbCrLf)
    s = Replace(s, vbCrLf, " ")
    For Each code In Array(13, 10, 11, 12, &H85, &H2028, &H2029)
        st.lineBreaks = st.lineBreaks + CountOf(s, ChrW(code))
        s = Replace(s, ChrW(code), " ")
    Next code
    FixLineBreaks = s
End Function

' Пробельные символы из веба, SAP и PDF, которые выглядят как обычный пробел (Unicode White_Space):
' табуляция, неразрывный (A0), огамский (1680), U+2000-200A, узкий неразрывный (202F),
' математический (205F), идеографический (3000). Нулевой ширины (200B-200D, 2060, FEFF) - удаляются.
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

' ---------------------------------------------------------------------------------------------
' Отчёт

' Пробел в начале ключа - всегда мусор, и глазами его не видно, поэтому края чистятся ВСЕГДА,
' какие бы флажки ни стояли (владелец, 2026-09-22). Исключение одно: если замена переносов строк
' выключена, перенос по краю не трогаем - его оставили намеренно.
Private Function TrimEdges(s As String, ByRef st As Stats, ByVal trimBreaks As Boolean) As String
    Dim i As Long, j As Long, k As Long, before As Long, res As String, keptHead As String, keptTail As String
    before = Len(s)
    ' Границы "пробельных" серий по краям. Обход идёт СКВОЗЬ серию: перенос строки, который велено
    ' сохранить, не должен прикрывать собой соседний обычный пробел (замечание Codex, 7-й круг).
    i = 1
    Do While i <= Len(s)
        If Not IsSpaceLike(Mid$(s, i, 1)) Then Exit Do
        i = i + 1
    Loop
    If i > Len(s) Then                     ' вся ячейка из пробельных
        If trimBreaks Then
            st.edgesTrimmed = st.edgesTrimmed + before
            Exit Function                  ' пусто
        End If
        For k = 1 To Len(s)
            If IsLineBreak(Mid$(s, k, 1)) Then keptHead = keptHead & Mid$(s, k, 1)
        Next k
        st.edgesTrimmed = st.edgesTrimmed + (before - Len(keptHead))
        TrimEdges = keptHead
        Exit Function
    End If
    j = Len(s)
    Do While j >= i
        If Not IsSpaceLike(Mid$(s, j, 1)) Then Exit Do
        j = j - 1
    Loop
    If Not trimBreaks Then                 ' переносы по краям остаются, всё остальное уходит
        For k = 1 To i - 1
            If IsLineBreak(Mid$(s, k, 1)) Then keptHead = keptHead & Mid$(s, k, 1)
        Next k
        For k = j + 1 To Len(s)
            If IsLineBreak(Mid$(s, k, 1)) Then keptTail = keptTail & Mid$(s, k, 1)
        Next k
    End If
    res = keptHead & Mid$(s, i, j - i + 1) & keptTail
    st.edgesTrimmed = st.edgesTrimmed + (before - Len(res))
    TrimEdges = res
End Function

Private Function IsLineBreak(ByVal ch As String) As Boolean
    Select Case AscW(ch) And &HFFFF&
        Case 10, 11, 12, 13, &H85, &H2028, &H2029
            IsLineBreak = True
    End Select
End Function

Private Function WhatWasThere(ByRef st As Stats) As String
    WhatWasThere = "В выделении: " & Format$(st.selected, "#,##0") & " ячеек - текстовых " & st.textCells & _
                   ", формул " & st.formulas & ", чисел и дат " & st.numbers & ", логических и ошибок " & st.others & _
                   ", пустых " & Format$(st.empties, "#,##0") & "."
End Function

Private Function ListAddresses(ByVal r As Range, ByVal maxListed As Long) As String
    Dim c As Range, k As Long, out As String
    For Each c In r
        k = k + 1
        If k <= maxListed Then out = out & IIf(k > 1, ", ", "") & c.Address(False, False)
    Next c
    If k > maxListed Then out = out & " и ещё " & (k - maxListed)
    ListAddresses = out
End Function

' Исключения идут ПЕРВЫМИ и в бюджет MsgBox укладываются раньше статистики: если что-то придётся
' обрезать, это будут цифры, а не адреса.
Private Function Report(ByRef st As Stats, ByVal c As Boolean, ByVal l As Boolean, ByVal r As Boolean, ByVal n As Boolean, _
                        ByVal onlyCodes As Boolean, ByVal flattened As Range, ByVal mergedPartial As Range, _
                        ByVal headers As Range, ByVal selectedOk As Boolean, ByVal backupLeftOpen As Boolean, _
                        ByVal backupName As String, ByVal appLeft As String, ByVal postErr As String, _
                        ByVal nothingChosen As Boolean) As String
    Dim head As String, alarm As String, exc As String, stat As String
    ' Бюджет MsgBox: аварийные предупреждения идут ПЕРВЫМИ и не обрезаются никогда; статистика
    ' добавляется, только если остаётся место. Адресов ячеек в отчёте нет (владелец, 2026-09-24):
    ' то, на что стоит посмотреть, выделено на листе, остальное - числом.
    alarm = ReportAlarm(backupLeftOpen, backupName, appLeft, postErr)
    head = ReportHead(st, nothingChosen)
    exc = ReportExceptions(st, onlyCodes, flattened, mergedPartial, headers, selectedOk)
    stat = ReportStats(st, c, l, r, n, onlyCodes, Len(exc) = 0)
    Report = head & alarm & exc
    If Len(Report) + Len(stat) <= REPORT_BUDGET Then Report = Report & stat
End Function
Private Function ReportAlarm(ByVal backupLeftOpen As Boolean, ByVal backupName As String, _
                             ByVal appLeft As String, ByVal postErr As String) As String
    Dim a As String
    ' Каждая часть ОГРАНИЧЕНА по длине. Описание ошибки от Excel бывает в тысячу символов, и без
    ' ограничения оно вытеснило бы из окна два следующих предупреждения - то есть ровно то, что
    ' пользователю надо сделать руками. Сумма частей заведомо меньше REPORT_BUDGET.
    If Len(postErr) > 0 Then
        a = a & vbCrLf & "ПОСЛЕ ЗАПИСИ произошла ошибка: " & Clip(postErr, ALARM_ERR_MAX) & vbCrLf & _
            "(ячейки уже очищены и перечитаны - откат не требовался)" & vbCrLf
    End If
    If backupLeftOpen Then
        a = a & vbCrLf & "Временную книгу " & Clip(backupName, ALARM_NAME_MAX) & " закрыть не удалось - закройте её вручную, НЕ сохраняя." & vbCrLf
    End If
    If Len(appLeft) > 0 Then a = a & vbCrLf & AppLeftText(Clip(appLeft, ALARM_STATE_MAX)) & vbCrLf
    ReportAlarm = a
End Function

Private Function Clip(ByVal s As String, ByVal n As Long) As String
    If Len(s) <= n Then Clip = s Else Clip = Left$(s, n - 3) & "..."
End Function

Private Function ReportHead(ByRef st As Stats, ByVal nothingOn As Boolean) As String
    Dim head As String
    head = "Обработка завершена." & IIf(nothingOn, " Ни одного действия не выбрано - очищены только края ячеек.", "") & _
           vbCrLf & "Изменено текстовых ячеек: " & st.cellsChanged & " из " & st.textCells
    If st.richKept + st.richFlattened > 0 Then head = head & " (с оформлением символов: " & st.richKept & " правлено посимвольно" & IIf(st.richFlattened > 0, ", " & st.richFlattened & " записано целиком", "") & ")"
    If st.hiddenText > 0 Then head = head & " (текстовых ячеек в скрытых строках/столбцах: " & st.hiddenText & ", обработаны наравне с видимыми)"
    ReportHead = head & vbCrLf
End Function

Private Function ReportExceptions(ByRef st As Stats, ByVal onlyCodes As Boolean, ByVal flattened As Range, _
                                  ByVal mergedPartial As Range, ByVal headers As Range, ByVal selectedOk As Boolean) As String
    Dim exc As String
    ' На что стоит посмотреть - числом здесь, а сами ячейки выделены на листе.
    If st.cyrMixed > 0 Then
        If onlyCodes Then
            exc = exc & vbCrLf & "БУКВЫ БЕЗ ЛАТИНСКОЙ ПАРЫ (Ж, Ш, У, Ы...) остались в " & st.cyrMixed & " яч." & vbCrLf
        Else
            exc = exc & vbCrLf & "ЛАТИНИЦА И КИРИЛЛИЦА В ОДНОМ СЛОВЕ: " & st.cyrMixed & " яч." & vbCrLf & _
                  "(похоже на код, но правило не решилось; если это коды - повторите на выделении с флажком 5)" & vbCrLf
        End If
    End If
    If Not headers Is Nothing Then
        exc = exc & vbCrLf & "НЕ ТРОНУТО " & st.tableHeaders & " заголовков таблиц Excel" & vbCrLf & _
              "(Excel сам переименовывает дубликаты и правит ссылки на столбец; переименуйте вручную)" & vbCrLf
    End If
    If Not mergedPartial Is Nothing Then
        exc = exc & vbCrLf & "НЕ ТРОНУТО " & st.mergedPartial & " объединённых яч., выделенных не целиком" & vbCrLf
    End If
    If Not flattened Is Nothing Then
        exc = exc & vbCrLf & "ЗАПИСАНЫ ЦЕЛИКОМ " & st.richFlattened & " яч. с оформлением символов" & vbCrLf & _
              "(посимвольная правка не сошлась; Excel оставляет оформление по позициям, и оно могло съехать - проверьте глазами)" & vbCrLf
    End If
    If Len(exc) > 0 Then
        exc = exc & IIf(selectedOk, "Эти ячейки сейчас ВЫДЕЛЕНЫ на листе.", _
                        "Выделить их на листе НЕ удалось (лист скрыт или выделение перехвачено).") & vbCrLf
    End If
    ' Русский текст - не исключение: его проверили и оставили намеренно. Только число, без выделения.
    If st.cyrLeft - st.cyrMixed > 0 Then
        exc = exc & vbCrLf & "Русский текст оставлен как есть: " & (st.cyrLeft - st.cyrMixed) & " яч. (так и должно быть, не выделяется)." & vbCrLf
    End If
    ' Обязательное, но не аварийное: ячейка обработана полностью, просто Excel повёл себя так, и
    ' молчать об этом нельзя.
    If st.crlfToLf > 0 Then
        exc = exc & vbCrLf & "CRLF -> LF в " & st.crlfToLf & " яч. с оформлением символов: Excel сам убирает CR при посимвольной правке." & vbCrLf & _
              "(перенос строки на месте, оформление сохранено - ячейки обработаны полностью)" & vbCrLf
    End If
    ReportExceptions = exc
End Function
Private Function ReportStats(ByRef st As Stats, ByVal c As Boolean, ByVal l As Boolean, ByVal r As Boolean, ByVal n As Boolean, _
                             ByVal onlyCodes As Boolean, ByVal noExceptions As Boolean) As String
    Dim stat As String
    stat = vbCrLf & WhatWasThere(st) & vbCrLf
    If c Then stat = stat & "- Кириллица: заменено символов " & st.homoglyphs & IIf(onlyCodes, " (везде, без проверки на код)", "") & vbCrLf
    If l Then stat = stat & "- Переносов строк заменено: " & st.lineBreaks & vbCrLf
    If r Then stat = stat & "- Пробельных и невидимых символов удалено: " & st.spacesRemoved & vbCrLf
    If n Then stat = stat & "- Пробельных и невидимых символов убрано при нормализации: " & st.spacesCollapsed & vbCrLf
    ' Строка есть ВСЕГДА, даже с нулём: очистка краёв не зависит от флажков, и отчёт должен это
    ' показывать. Если края уже убрали действия 3 или 4, эти символы посчитаны в их строках выше.
    stat = stat & "- Пробельных и невидимых символов убрано по краям: " & st.edgesTrimmed & " (края чистятся всегда)" & vbCrLf
    If st.cellsChanged = 0 And noExceptions Then stat = stat & vbCrLf & "Изменений не найдено."
    ReportStats = stat
End Function
