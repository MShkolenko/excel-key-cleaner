Attribute VB_Name = "Ribbon"
Option Explicit

' Кнопки надстройки excel-key-cleaner.xlam: вкладка "Очистка ключей" на ленте и пункт в меню
' правой кнопки по ячейке. Разметка - customUI/customUI14.xml внутри .xlam (исходник -
' src/customUI14.xml), onAction там указывает сюда. Сам макрос - CleanKeys в модуле Cleaning.

Private ui As IRibbonUI

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set ui = ribbon
End Sub

' Для проверки сборки (tests/verify_xlam.py): Excel разобрал разметку ленты и вызвал onLoad,
' и вкладку можно показать - так снимается картинка ленты для README.
Public Function RibbonLoaded() As Boolean
    RibbonLoaded = Not ui Is Nothing
End Function

Public Sub ShowTab()
    If Not ui Is Nothing Then ui.ActivateTab "tabKeyCleaner"
End Sub

Public Sub Ribbon_Clean(control As IRibbonControl)
    Cleaning.CleanKeys          ' явно свой модуль: CleanKeys может быть ещё и в PERSONAL.XLSB
End Sub

Public Sub Ribbon_Help(control As IRibbonControl)
    ThisWorkbook.FollowHyperlink "https://github.com/MShkolenko/excel-key-cleaner#readme"
End Sub
