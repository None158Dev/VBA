VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HandlerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private shHadler As Worksheet
Private NumberError%
'==================================================================================================================
Private shall As Sheets
Private ArrTotal As Variant
Private ArrRange As Variant
Private Sub Class_Initialize()
Set shall = ThisWorkbook.Sheets
ReDim ArrTotal(shall.Count - 3, 1 To 8)
ReDim ArrRange(shall.Count - 3)
End Sub
Public Sub StartHandler()
Call Total
If ErrorBoolean Then PrintError = HadlerError(NumberError): Exit Sub
End Sub
Private Function Total()
Dim LBoundRow&, UBoundRow&
Dim lastBorder As Boolean
Dim CheckSection As Object
Dim CheckColPrice As Object
Dim CheckNumbSmeta As Object
Dim CheckToWhat As Object
Dim CheckTotal As Object

For Each shHadler In shall
If shHadler.CodeName <> "Data" And shHadler.CodeName <> "Statistic" Then
    UBoundRow = shHadler.Cells(Rows.Count, 1).End(xlUp).Row
        Do: UBoundRow = UBoundRow + 1
            If shHadler.Cells(UBoundRow, 1).Borders(xlEdgeBottom).LineStyle <> 1 Then lastBorder = True
        Loop Until lastBorder: UBoundRow = UBoundRow - 1
'==================================================================================================================
        Set CheckSection = shHadler.Range("A1:O" & UBoundRow).Find("*Раздел*")
        Set CheckColPrice = shHadler.Range("A1:O" & UBoundRow).Find("*Всего в текущих (прогнозных) ценах*")
        Set CheckNumbSmeta = shHadler.Range("A1:O" & UBoundRow).Find("*локальный сметный расчет*")
        Set CheckToWhat = shHadler.Range("A1:O" & UBoundRow).Find("*наименование работ и затрат, наименование объекта*")
        Set CheckTotal = shHadler.Range("A1:O" & UBoundRow).Find("*в текущем уровне цен*")
        'проверки сметы на правильность составления
        If CheckSection Is Nothing Then ErrorBoolean = True: NumberError = 1: Exit Function
        If Not CheckSection Is Nothing Then LBoundRow = CheckSection.Row
        If LBoundRow > 50 Then ErrorBoolean = True: NumberError = 2: Exit Function
        If CheckColPrice.Column <> 10 Then ErrorBoolean = True: NumberError = 3: Exit Function
        If CheckNumbSmeta Is Nothing Then ErrorBoolean = True: NumberError = 4: Exit Function
        If CheckToWhat Is Nothing Then ErrorBoolean = True: NumberError = 5: Exit Function
        If CheckTotal Is Nothing Then ErrorBoolean = True: NumberError = 6: Exit Function
        
        ArrRange(shHadler.index - 3) = shHadler.Range("A" & LBoundRow & ":K" & UBoundRow)
        ArrTotal(shHadler.index - 3, 1) = shHadler.index - 2
        ArrTotal(shHadler.index - 3, 2) = Clear(CheckNumbSmeta.Offset(-1), 1)
        ArrTotal(shHadler.index - 3, 3) = Clear(CheckTotal.Offset(1), 2)
        ArrTotal(shHadler.index - 3, 6) = Clear(CheckToWhat.Offset(-1), 3)
    End If
Next shHadler

End Function
Private Function HadlerError(ByVal NumbError As Integer) As String
If NumbError = 1 Then HadlerError = "Нет РАЗДЕЛА на ЛИСТЕ " & Chr(171) & shHadler.Name & Chr(187)
If NumbError = 2 Then HadlerError = "Нет РАЗДЕЛА в начале СМЕТЫ на ЛИСТЕ " & Chr(171) & shHadler.Name & Chr(187)
If NumbError = 3 Then HadlerError = "Наименование " & Chr(171) & "ВСЕГО В ТЕКУЩИХ (ПРОГНОЗНЫХ) ЦЕНАХ, РУБ." & Chr(187) & " находится не в 10 СТОЛБЦЕ на ЛИСТЕ " & Chr(171) & shHadler.Name & Chr(187)
If NumbError = 4 Then HadlerError = "В смете нет СТРОКИ " & Chr(171) & "ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ" & Chr(187) & " на ЛИСТЕ " & Chr(171) & shHadler.Name & Chr(187)
If NumbError = 5 Then HadlerError = "В смете нет СТРОКИ " & Chr(171) & "НАИМЕНОВАНИЕ РАБОТ И ЗАТРАТ, НАИМЕНОВАНИЕ ОБЪЕКТА" & Chr(187) & " на ЛИСТЕ " & Chr(171) & shHadler.Name & Chr(187)
If NumbError = 6 Then HadlerError = "В смете нет СТРОКИ " & Chr(171) & "В ТЕКУЩЕМ УРОВНЕ ЦЕН" & Chr(187) & " на ЛИСТЕ " & Chr(171) & shHadler.Name & Chr(187)
End Function
Private Function Clear(ByVal value As Variant, ByVal Numb As Integer) As Variant
If Numb = 1 Then Clear = CStr(Application.WorksheetFunction.Trim(Replace(Split(value, "№")(1), "-", "_")))
If Numb = 2 Then Clear = CDbl(Application.WorksheetFunction.Trim(Split(value, " ")(0)))
If Numb = 3 Then Clear = CStr(Application.WorksheetFunction.Trim(Replace(value, "на ", "")))
End Function
Public Property Get ArrSvod()
ArrSvod = ArrTotal
End Property
Public Property Get ArrSmets()
ArrSmets = ArrRange
End Property

