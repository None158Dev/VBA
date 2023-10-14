Attribute VB_Name = "DataGroup"
Option Explicit
Private Sub Group()
Dim Infinity As Double
Dim i&
Dim clsDict As MyGroup
Set clsDict = New MyGroup

Call DisabledApps(False, False)
Infinity = Timer

For i = 1 To 4
    On Error Resume Next: [BopSebes].Rows.Ungroup: On Error GoTo 0
Next i

clsDict.Start

Call DisabledApps(True, True)
FILE_ControlPanel.ComplitTimeGroup.Visible = True: FILE_ControlPanel.ComplitTimeGroup.Caption = "Готово! Затрачено времени: " & Format(Timer - Infinity, "0.00 сек")


End Sub
