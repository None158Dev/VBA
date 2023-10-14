Attribute VB_Name = "DataKeysUniq"
Option Explicit
Private Sub SmUniqKeys()
Dim clsDict As MyKeysUniq
Dim Infinity As Double
Set clsDict = New MyKeysUniq

Call DisabledApps(False, False)
Infinity = Timer

clsDict.Start

On Error Resume Next: [Coefficient].Delete: On Error GoTo 0
On Error Resume Next: [CountSm].Delete: On Error GoTo 0

[Coefficient].Resize(UBound(clsDict.UniqKeys), 5) = clsDict.UniqKeys
[CountSm].Resize(UBound(clsDict.UniqKeys), 2) = clsDict.UniqKeys

Call DisabledApps(True, True)
FILE_ControlPanel.ComplitTimeSvod.Visible = True: FILE_ControlPanel.ComplitTimeSvod.Caption = "Готово! Затрачено времени: " & Format(Timer - Infinity, "0.00 сек")
End Sub
