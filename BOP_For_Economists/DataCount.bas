Attribute VB_Name = "DataCount"
Option Explicit
Private Sub CountPetcent()
Dim Infinity As Double
Dim clsDict As MyCount

Call DisabledApps(False, False)
Infinity = Timer
[pos_UniqP].Columns(14).Calculate
[pos_UniqM].Columns(14).Calculate
Set clsDict = New MyCount

clsDict.CountCompAll
clsDict.CountCompSM

[CountPM].Columns(2).Resize(UBound(clsDict.UniqCompPM), 5) = clsDict.UniqCompPM
[CountSm].Columns(3).Resize(UBound(clsDict.UniqCompSM), 9) = clsDict.UniqCompSM

Call DisabledApps(True, True)
FILE_ControlPanel.ComplitTimePercent.Visible = True: FILE_ControlPanel.ComplitTimePercent.Caption = "Готово! Затрачено времени: " & Format(Timer - Infinity, "0.00 сек")

End Sub
