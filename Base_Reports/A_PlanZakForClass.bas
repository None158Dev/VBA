Attribute VB_Name = "A_PlanZakForClass"
Option Explicit
Private Sub DataMain()
Dim Infinity As Double
Dim clsDic As PlanZakClass
Set clsDic = New PlanZakClass

If [ForPlanZak].ListObject.DataBodyRange Is Nothing Then MsgBox "Ключей на листе " & Chr(171) & "ОТЧЁТ" & Chr(187) & " нет!", vbCritical: Exit Sub

Call DisabledApps(False, False)
Infinity = Timer

clsDic.Start

On Error Resume Next: [PlanZak].Delete: On Error GoTo 0

[PlanZak].Resize(UBound(clsDic.ItemsRes), 22) = clsDic.ItemsRes

[PlanZak].Worksheet.Visible = True
Call FilterColor
Call CapDate
Call DisabledApps(True, True)
FILE_FRM_Search.ComplitTimePlanZak.Visible = True: FILE_FRM_Search.ComplitTimePlanZak.Caption = "Готово! Затрачено времени:" & Chr(11) & Format(Timer - Infinity, "0.00 сек")

End Sub
Private Sub FilterColor()
Dim shSeb As ListObject
Dim i%
Dim arr As Variant
Set shSeb = [PlanZak].Worksheet.ListObjects(1)
arr = Array(12, 13, 14, 16, 17, 18, 20, 21, 22)
    
    shSeb.ListRows.Add (1): [PlanZak].Columns(6).Rows(1) = "Total"
    [PlanZak].Interior.ColorIndex = xlNone: [PlanZak].Rows(1).Interior.Color = 65535
    [PlanZak].AutoFilter Field:=6, Criteria1:="Смета", Operator:=xlFilterValues: [PlanZak].SpecialCells(xlCellTypeVisible).Interior.Color = 16776960
    
    For i = LBound(arr) To UBound(arr)
        [PlanZak].Columns(arr(i)).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=SUMIFS(C,C2,[@[Ключ сметы]],C6,""<>Смета"")"
        [PlanZak].Columns(arr(i)).Rows(1).FormulaR1C1 = "=SUMIFS(C,C6,""Смета"")"
    Next i

shSeb.AutoFilter.ShowAllData
End Sub
Private Sub CapDate()
Dim i&, col%
Dim MonthText As String
Dim YearNumb As String
Dim QuarterNumb As String

For i = 0 To 2
    MonthText = MonthName(Month([ForPlanZak].Columns(2 + i).Rows(0)))
    YearNumb = Year([ForPlanZak].Columns(2 + i).Rows(0))
    QuarterNumb = DatePart("q", [ForPlanZak].Columns(2 + i).Rows(0))
    
    [PlanZak].Columns(11 + col).Rows(0) = "Кол-во " & MonthText & " " & YearNumb
    [PlanZak].Columns(12 + col).Rows(0) = "Сумма заказчика " & MonthText & " " & YearNumb
    [PlanZak].Columns(13 + col).Rows(0) = "Сумма поставки " & MonthText & " " & YearNumb
    [PlanZak].Columns(14 + col).Rows(0) = "Сумма оплаты " & MonthText & " " & YearNumb
    col = col + 4
Next i

[PlanZak].Worksheet.Cells(3, 6) = "Помесячный план закупки материалов в натуральных единицах (без разбивки по поставщикам) на " & QuarterNumb & _
              " квартал " & YearNumb & " года, необходимых для выполнения графика производственных работ, согласно финансовой " & _
              "модели по строительству комплекса зданий, строений, сооружений КФ МГТУ им. Н. Э. Баумана"

End Sub
