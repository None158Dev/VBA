Attribute VB_Name = "A_Function"
Option Explicit
Private Function SaveList(Optional Path As String)
Path = ExistsFolder

If FILE_FRM_Search.Job.Value = True Then
    If [pos_P].Worksheet.Visible = False Then MsgBox "Нет сформированного отчёта уникальных работ!", vbCritical: Exit Function
    [pos_P].Worksheet.Copy
    With ActiveWorkbook
        .SaveAs Path & [pos_P].Worksheet.Name, FileFormat:=xlExcel12
        .Worksheets(1).Shapes("SearchP").OnAction = "Р.SearchPosP"
        .Worksheets(1).Shapes("import").OnAction = "Р.importfile"
        .Close SaveChanges:=True
    End With
    On Error Resume Next: [pos_P].Delete: On Error GoTo 0
    [pos_P].Worksheet.Visible = False
End If

If FILE_FRM_Search.Material.Value = True Then
    If [pos_M].Worksheet.Visible = False Then MsgBox "Нет сформированного отчёта уникальных материалов!", vbCritical: Exit Function
    [pos_M].Worksheet.Copy
    With ActiveWorkbook
        .SaveAs Path & [pos_M].Worksheet.Name, FileFormat:=xlExcel12
        .Worksheets(1).Shapes("SearchM").OnAction = "М.SearchPosM"
        .Worksheets(1).Shapes("import").OnAction = "М.importfile"
        .Close SaveChanges:=True
    End With
    On Error Resume Next: [pos_M].Delete: On Error GoTo 0
    [pos_M].Worksheet.Visible = False
End If

If FILE_FRM_Search.PlanZak.Value = True Then
    If [PlanZak].Worksheet.Visible = False Then MsgBox "Нет сформированного отчёта плана закупок!", vbCritical: Exit Function
    [PlanZak].Worksheet.Copy
    With ActiveWorkbook
        .SaveAs Path & [PlanZak].Worksheet.Name, FileFormat:=xlExcel12
        .Close SaveChanges:=True
    End With
    On Error Resume Next: [PlanZak].Delete: On Error GoTo 0
    [PlanZak].Worksheet.Visible = False
End If

If FILE_FRM_Search.Bop.Value = True Then
    If [Bop].Worksheet.Visible = False Then MsgBox "Нет сформированного отчёта ведомости объёмов работ!", vbCritical: Exit Function
    On Error Resume Next: [Bop].Delete: On Error GoTo 0
    [Bop].Worksheet.Visible = False
End If

If FILE_FRM_Search.Sebes.Value = True Then
    If [Sebes].Worksheet.Visible = False Then MsgBox "Нет сформированного отчёта себестоимости!", vbCritical: Exit Function
    On Error Resume Next: [Sebes].Delete: On Error GoTo 0
    [Sebes].Worksheet.Visible = False
End If

If FILE_FRM_Search.Cntr.Value = True Then
    If shCntr Is Nothing Then MsgBox "Нет сформированной контрактации!", vbCritical: Exit Function
    Application.DisplayAlerts = False
    shCntr.Copy
    With ActiveWorkbook
        .SaveAs Path & "НРВ " & Split(shCntr.Cells(4, 8), ":")(0) & "_" & Replace(Split(shCntr.Cells(4, 8), "«")(1), "»", ""), FileFormat:=xlExcel12
        .Close SaveChanges:=True
    End With
    On Error Resume Next: shCntr.Delete: On Error GoTo 0
    Application.DisplayAlerts = True
End If

[est].Worksheet.Activate

End Function
Private Function ExistsFolder() As String
Dim FileSystemObject As New Scripting.FileSystemObject

If FileSystemObject.FolderExists(ThisWorkbook.Path & "\Для отчётов") Then ExistsFolder = ThisWorkbook.Path & "\Для отчётов" & "\": Exit Function
If Dir(ThisWorkbook.Path & "\Для отчётов") = "" Then CreateObject("Scripting.FileSystemObject").CreateFolder (ThisWorkbook.Path & "\Для отчётов"): ExistsFolder = ThisWorkbook.Path & "\Для отчётов" & "\": Exit Function

End Function
Private Function shCntr(Optional i As Integer) As Object
For i = 1 To ThisWorkbook.Worksheets.Count
    With ThisWorkbook
        If .Worksheets(i).Name Like "КОНТРАКТАЦИЯ*" Then
            Set shCntr = .Worksheets(i)
        End If
    End With
Next i
End Function
Public Function DisabledApps(OnOff As Boolean, CalculateApp As Boolean): If CalculateApp = False Then CalculateApp = xlManual Else CalculateApp = xlAutomatic
Application.ScreenUpdating = OnOff
Application.Calculation = CalculateApp
Application.AskToUpdateLinks = OnOff
Application.DisplayAlerts = OnOff
End Function
Private Sub UpdateTable()
Application.Run "shCntr.Cntr_ReFresh"
Application.Run "shEst.Est_ReFresh"
Application.Run "shPart.Part_ReFresh"
Application.Run "shPos.Pos_ReFresh"
Application.Run "shWO.WO_ReFresh"
End Sub

