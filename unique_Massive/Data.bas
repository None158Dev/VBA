Attribute VB_Name = "Data"
Option Explicit
Public Sub DataMain()
Dim Apps As Variant
Dim Infinity As Double
Dim clsDic As MyDictionary
Set clsDic = New MyDictionary

If FILE_ControlPanel.MinUniq.Value = False And FILE_ControlPanel.MaxUniq.Value = False Then FILE_ControlPanel.UniqFalse.Visible = True: Exit Sub
If FILE_ControlPanel.MinUniq.Value = True Or FILE_ControlPanel.MaxUniq.Value = True Then FILE_ControlPanel.UniqFalse.Visible = False

Call DisabledApps(False, False)
Infinity = Timer

clsDic.add [pos_all].Value2

On Error Resume Next: [pos_P].Delete
On Error Resume Next: [pos_M].Delete
[pos_P].Resize(UBound(clsDic.ItemsP), 10) = clsDic.ItemsP
[pos_M].Resize(UBound(clsDic.ItemsM), 10) = clsDic.ItemsM
Call DisabledApps(True, True)
FILE_ControlPanel.ComplitTime.Visible = True: FILE_ControlPanel.ComplitTime.Caption = "Готово! Затрачено времени:" & Chr(11) & Format(Timer - Infinity, "0.00 сек")
End Sub
Private Function SaveList(Optional Path As String)
Path = [pos_all].Worksheet.Range("XFD1") & "\"

If FILE_ControlPanel.Job.Value = False And FILE_ControlPanel.Material.Value = False Then FILE_ControlPanel.SaveFalse.Visible = True
If FILE_ControlPanel.Job.Value = True Or FILE_ControlPanel.Material.Value = True Then FILE_ControlPanel.SaveFalse.Visible = False

If FILE_ControlPanel.Job.Value = True Then
    [pos_P].Worksheet.Copy
    With ActiveWorkbook
        .SaveAs Path & [pos_P].Worksheet.Name, FileFormat:=xlExcel12
        .Worksheets(1).Shapes("SearchP").OnAction = "Р.SearchPosP"
        .Worksheets(1).Shapes("import").OnAction = "Р.importfile"
        .Close SaveChanges:=True
    End With
End If

If FILE_ControlPanel.Material.Value = True Then
    [pos_M].Worksheet.Copy
    With ActiveWorkbook
        .SaveAs Path & [pos_M].Worksheet.Name, FileFormat:=xlExcel12
        .Worksheets(1).Shapes("SearchM").OnAction = "М.SearchPosM"
        .Worksheets(1).Shapes("import").OnAction = "М.importfile"
        .Close SaveChanges:=True
    End With
End If
End Function
Public Function DisabledApps(OnOff As Boolean, CalculateApp As Boolean): If CalculateApp = False Then CalculateApp = xlManual Else CalculateApp = xlAutomatic
Application.ScreenUpdating = OnOff
Application.Calculation = CalculateApp
Application.AskToUpdateLinks = OnOff
Application.DisplayAlerts = OnOff
End Function



