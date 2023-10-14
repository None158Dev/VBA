Attribute VB_Name = "KeysPos"
Option Explicit
Private Sub Main()
Dim Infinity As Double
Dim clsKeyGen As New KeyGenClass
Set clsKeyGen = New KeyGenClass
Infinity = Timer

clsKeyGen.KeysPosStart

[DataRes].Columns(7) = clsKeyGen.arrResultsKey
[DataRes].Columns(11) = clsKeyGen.arrResultsKeyPos
[DataTotal].Columns(7).Resize(, 2) = clsKeyGen.arrResultsTotal
FILE_ControlPanel.ComplitTimeKeys.Visible = True: FILE_ControlPanel.ComplitTimeKeys.Caption = "Готово! Затрачено времени:" & Chr(11) & Format(Timer - Infinity, "0.00 сек")
End Sub
