Attribute VB_Name = "Func"
Option Explicit
Sub importfile()              'Сбор всех смет из папки
    Dim fnameList, fnameCurFile As Variant
    Dim countFiles, countSheets As Integer
    Dim wksCurSheet As Worksheet
    Dim wbkCurBook, wbkSrcBook As Workbook
 
    fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", Title:="Выберите файлы для объединения", MultiSelect:=True)
 
    If (vbBoolean <> VarType(fnameList)) Then
 
        If (UBound(fnameList) > 0) Then
            countFiles = 0
            countSheets = 0
 
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
 
            Set wbkCurBook = ActiveWorkbook
 
            For Each fnameCurFile In fnameList
                countFiles = countFiles + 1
 
                Set wbkSrcBook = Workbooks.Open(Filename:=fnameCurFile)
 
                For Each wksCurSheet In wbkSrcBook.Sheets
                    countSheets = countSheets + 1
                    wksCurSheet.Copy after:=wbkCurBook.Sheets(wbkCurBook.Sheets.Count)
                Next
 
                wbkSrcBook.Close SaveChanges:=False
 
            Next
 
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Worksheets("Итоги после обработки").Activate
            
            MsgBox "Из " & countFiles & " файлов" & vbCrLf & "Собрано " & countSheets & " листа", Title:="Объединение файлов"
        End If
 
    Else
        MsgBox "Нет выбранных файлов!", vbCritical, "Объединение файлов"
    End If
End Sub
Sub Sheets_Del()
Dim i&
If Sheets.Count <> 2 Then
Application.DisplayAlerts = False

    For i = Sheets.Count To 3 Step -1
            Worksheets(i).Delete
    Next
Else
    MsgBox "Нет смет, которые можно удалить!", vbCritical
End If
Application.DisplayAlerts = True

End Sub
Public Function DisabledApps(OnOff As Boolean, CalculateApp As Boolean): If CalculateApp = False Then CalculateApp = xlManual Else CalculateApp = xlAutomatic
Application.ScreenUpdating = OnOff
Application.Calculation = CalculateApp
Application.AskToUpdateLinks = OnOff
Application.DisplayAlerts = OnOff
End Function

