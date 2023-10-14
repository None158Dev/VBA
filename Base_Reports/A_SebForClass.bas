Attribute VB_Name = "A_SebForClass"
Option Explicit
Private Sub DataMain()
Dim Infinity As Double
Dim clsDic As SebClass
Set clsDic = New SebClass

If [SebesCoeff].ListObject.DataBodyRange Is Nothing Then MsgBox "Ключей на листе " & Chr(171) & "ОТЧЁТ" & Chr(187) & " нет!", vbCritical: Exit Sub

Call DisabledApps(False, False)
Infinity = Timer

clsDic.Start

On Error Resume Next: [Sebes].Delete: On Error GoTo 0
[Sebes].Resize(UBound(clsDic.ItemsRes), 23) = clsDic.ItemsRes
[Sebes].Worksheet.Visible = True
Call DisabledApps(True, True)

FILE_FRM_Search.ComplitTimeSebes.Visible = True: FILE_FRM_Search.ComplitTimeSebes.Caption = "Готово! Затрачено времени:" & Chr(11) & Format(Timer - Infinity, "0.00 сек")

End Sub
