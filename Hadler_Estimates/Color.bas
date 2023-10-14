Attribute VB_Name = "Color"
Option Explicit
Private Sub Main()
Dim Infinity As Double
Dim clsHandler As New ColorClass
Infinity = Timer
Call DisabledApps(False, False)
Dim clsColor As New ColorClass
Set clsColor = New ColorClass

clsColor.StartColorClass


Call DisabledApps(True, True)
FILE_ControlPanel.ComplitTimeColor.Visible = True: FILE_ControlPanel.ComplitTimeColor.Caption = "Готово! Затрачено времени:" & Chr(11) & Format(Timer - Infinity, "0.00 сек")

End Sub
