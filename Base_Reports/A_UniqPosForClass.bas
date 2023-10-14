Attribute VB_Name = "A_UniqPosForClass"
Option Explicit
Private Sub DataMain()
Dim Infinity As Double
Dim clsDic As UniqPosClass
Set clsDic = New UniqPosClass

If FILE_FRM_Search.MinUniq.Value = False And FILE_FRM_Search.MaxUniq.Value = False Then FILE_FRM_Search.UniqFalse.Visible = True: Exit Sub
If FILE_FRM_Search.MinUniq.Value = True Or FILE_FRM_Search.MaxUniq.Value = True Then FILE_FRM_Search.UniqFalse.Visible = False

Call DisabledApps(False, False)
Infinity = Timer

clsDic.Start [pos].Value2

On Error Resume Next: [pos_P].Delete: On Error GoTo 0
On Error Resume Next: [pos_M].Delete: On Error GoTo 0
[pos_P].Worksheet.Visible = True
[pos_M].Worksheet.Visible = True
[pos_P].Resize(UBound(clsDic.ItemsP), 10) = clsDic.ItemsP
[pos_M].Resize(UBound(clsDic.ItemsM), 10) = clsDic.ItemsM
Call DisabledApps(True, True)
FILE_FRM_Search.ComplitTimeUniq.Visible = True: FILE_FRM_Search.ComplitTimeUniq.Caption = "Готово! Затрачено времени:" & Chr(11) & Format(Timer - Infinity, "0.00 сек")
End Sub



