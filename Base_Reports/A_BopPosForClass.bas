Attribute VB_Name = "A_BopPosForClass"
Option Explicit
Private Sub DataMain()
Dim Infinity As Double
Dim clsDic As BopClass
Set clsDic = New BopClass

Call DisabledApps(False, False)
Infinity = Timer

clsDic.Start

On Error Resume Next: [CalcSebes].Delete: On Error GoTo 0
[Bop].Resize(UBound(clsDic.ItemsRes), 13) = clsDic.ItemsRes
[Bop].Worksheet.Visible = True
Call DisabledApps(True, True)
FILE_FRM_Search.ComplitTimeBop.Visible = True: FILE_FRM_Search.ComplitTimeBop.Caption = "Готово! Затрачено времени:" & Chr(11) & Format(Timer - Infinity, "0.00 сек")


End Sub
