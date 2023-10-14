Attribute VB_Name = "Func"
Option Explicit
Public Function DisabledApps(OnOff As Boolean, CalculateApp As Boolean): If CalculateApp = False Then CalculateApp = xlManual Else CalculateApp = xlAutomatic
Application.ScreenUpdating = OnOff
Application.Calculation = CalculateApp
Application.AskToUpdateLinks = OnOff
Application.DisplayAlerts = OnOff
End Function
Private Sub FilterColorFormula()
Dim shSeb As ListObject
Dim Infinity As Double
Dim ColorStyle As Object

Infinity = Timer
Call DisabledApps(False, False)

For Each ColorStyle In [_Color]
    [BopSebes].AutoFilter Field:=7, Criteria1:=ColorStyle.Value2, Operator:=xlFilterValues
    [BopSebes].SpecialCells(xlCellTypeVisible).Interior.Color = ColorStyle.Interior.Color
    [BopSebes].Columns(13).SpecialCells(xlCellTypeVisible).Formula2R1C1 = "=IF([@Тип]=""Смета"",SUM(SUMIFS(C,C[-12],[@[№ Сметы]],C[-6],{""р"";""м"";""о""})),IF([@Тип]=""Раздел"",SUM(SUMIFS(C,C[-11],[@[Ключ раздела]],C[-6],{""р"";""м"";""о""})),IF([@Тип]=""Группа"",SUM(SUMIFS(C,C[-10],[@[Ключ группы]],C[-6],{""р"";""м"";""о""})))))"
Next ColorStyle

[BopSebes].AutoFilter Field:=7, Criteria1:="р", Operator:=xlFilterValues
[BopSebes].Columns(1).Resize(, 10).SpecialCells(xlCellTypeVisible).Interior.Color = 13431551

Set shSeb = [BopSebes].Worksheet.ListObjects(1): shSeb.AutoFilter.ShowAllData
Call DisabledApps(True, True)
FILE_ControlPanel.ComplitTimePaint.Visible = True: FILE_ControlPanel.ComplitTimePaint.Caption = "Готово! Затрачено времени:" & Format(Timer - Infinity, "0.00 сек")

End Sub
