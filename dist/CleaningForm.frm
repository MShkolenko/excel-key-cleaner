VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CleaningForm 
   Caption         =   "Очистка ключей"
   ClientHeight    =   3520
   ClientLeft      =   -460
   ClientTop       =   -2050
   ClientWidth     =   5510
   OleObjectBlob   =   "CleaningForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CleaningForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Подписи без спецсимволов: MSForms не подставляет глифы из других шрифтов, отсутствующий
' символ рисуется квадратом (вчерашние "?" на кнопках - та же беда на уровне кодировки).
' Макрос продолжает работу ТОЛЬКО если Tag = "OK" - это ставит одна кнопка "Выполнить".
' Крестик, Esc и "Отмена" оставляют Tag пустым/Cancel, и ничего не меняется.

Private Sub CommandButton1_Click()
    Me.Tag = "OK"
    Me.Hide
End Sub

Private Sub CommandButton2_Click()
    Me.Tag = "Cancel"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Tag = "Cancel"
        Me.Hide
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.Width = 600
    Me.Height = 450
    Me.Caption = "Очистка ключей"

    With Me.CheckBox1
        .Left = 30: .Top = 30: .Width = 540: .Height = 35: .Font.Size = 14
        .Caption = "1. Кириллица -> латиница (по умолчанию - только в кодах, см. п.5)"
        .Value = True
    End With
    With Me.CheckBox2
        .Left = 30: .Top = 75: .Width = 540: .Height = 35: .Font.Size = 14
        .Caption = "2. Переносы строк -> пробел"
        .Value = True
    End With
    With Me.CheckBox3
        .Left = 30: .Top = 120: .Width = 540: .Height = 35: .Font.Size = 14
        .Caption = "3. Удалить все пробелы (и табуляцию, неразрывные, невидимые)"
        .Value = False
    End With
    With Me.CheckBox4
        .Left = 30: .Top = 165: .Width = 540: .Height = 35: .Font.Size = 14
        .Caption = "4. Нормализация пробелов: один вместо нескольких, убрать невидимые"
        .Value = True
    End With

    With Me.Label1
        .Left = 30: .Top = 205: .Width = 540: .Height = 25
        .Font.Size = 12: .Font.Bold = True: .ForeColor = RGB(255, 0, 0)
        .Caption = "Действия 3 и 4 несовместимы. Пробелы по краям убираются ВСЕГДА"
    End With

    ' Пятый флажок добавляется на лету: элементы формы лежат в бинарном .frx, а его текстом
    ' не правят. Читается из макроса как CleaningForm.Controls("CheckBox5").
    Dim cb5 As Object
    Set cb5 = Me.Controls.Add("Forms.CheckBox.1", "CheckBox5", True)
    With cb5
        .Left = 30: .Top = 240: .Width = 540: .Height = 35: .Font.Size = 14
        .Caption = "5. В выделении ТОЛЬКО коды: кириллицу менять везде, без проверки"
        .Value = False
    End With

    ' Порядок Tab: флажки 1-5, потом кнопки; начальный фокус - на первом флажке.
    Me.CheckBox1.TabIndex = 0: Me.CheckBox2.TabIndex = 1: Me.CheckBox3.TabIndex = 2: Me.CheckBox4.TabIndex = 3
    cb5.TabIndex = 4: Me.CommandButton1.TabIndex = 5: Me.CommandButton2.TabIndex = 6

    With Me.CommandButton1
        .Caption = "Выполнить"
        .Left = 130: .Top = 300: .Width = 160: .Height = 50
        .Font.Size = 14: .Font.Bold = True
        .Default = True
    End With
    With Me.CommandButton2
        .Caption = "Отмена"
        .Left = 330: .Top = 300: .Width = 160: .Height = 50
        .Font.Size = 14: .Font.Bold = True
        .Cancel = True
    End With
End Sub

Private Sub CheckBox3_Click()
    If Me.CheckBox3.Value = True Then Me.CheckBox4.Value = False
End Sub

Private Sub CheckBox4_Click()
    If Me.CheckBox4.Value = True Then Me.CheckBox3.Value = False
End Sub
