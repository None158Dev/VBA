Attribute VB_Name = "DataUniq"
Option Explicit
Private Sub Uniq()
Dim Infinity As Double
Dim shP As ListObject
Dim shM As ListObject
Dim clsDict As MyUniq
Set clsDict = New MyUniq


If FILE_ControlPanel.MinUniq.Value = False And FILE_ControlPanel.MaxUniq.Value = False Then FILE_ControlPanel.UniqFalse.Visible = True: Exit Sub
If FILE_ControlPanel.MinUniq.Value = True Or FILE_ControlPanel.MaxUniq.Value = True Then FILE_ControlPanel.UniqFalse.Visible = False


Call DisabledApps(False, False)
Set shP = [pos_UniqP].Worksheet.ListObjects(1): shP.AutoFilter.ShowAllData
Set shM = [pos_UniqM].Worksheet.ListObjects(1): shM.AutoFilter.ShowAllData
Infinity = Timer
clsDict.Start [BopSebes].Value2, [pos_UniqP].Value2, [pos_UniqM].Value2

On Error Resume Next: [pos_UniqP].Delete: On Error GoTo 0
On Error Resume Next: [pos_UniqM].Delete: On Error GoTo 0

[pos_UniqP].Resize(UBound(clsDict.ItemsP), 13) = clsDict.ItemsP
[pos_UniqM].Resize(UBound(clsDict.ItemsM), 13) = clsDict.ItemsM

Call DisabledApps(True, True)
FILE_ControlPanel.ComplitTime.Visible = True: FILE_ControlPanel.ComplitTime.Caption = "Готово! Затрачено времени:" & Chr(11) & Format(Timer - Infinity, "0.00 сек")

End Sub
