Attribute VB_Name = "Handler"
Option Explicit
Public PrintError As String
Public ErrorBoolean As Boolean
Private Sub Main()
Dim Infinity As Double
Dim clsHandler As New HandlerClass
Infinity = Timer
Dim clsHandlerArray As New HandlerArrayClass

Set clsHandler = New HandlerClass
Set clsHandlerArray = New HandlerArrayClass

clsHandler.StartHandler
If ErrorBoolean Then MsgBox PrintError, vbCritical, "You shall not pass": ErrorBoolean = False: Exit Sub

clsHandlerArray.StartHandlerArray clsHandler.ArrSmets, clsHandler.ArrSvod

On Error Resume Next: [DataRes].Delete: On Error GoTo 0
On Error Resume Next: [DataTotal].Delete: On Error GoTo 0
On Error Resume Next: [DataUniq].Delete: On Error GoTo 0
[DataRes].ListObject.Resize [DataRes].ListObject.Range.Resize(UBound(clsHandlerArray.arrResultsMain) + 3): [DataRes] = clsHandlerArray.arrResultsMain
[DataTotal].ListObject.Resize [DataTotal].ListObject.Range.Resize(UBound(clsHandlerArray.arrResultsTotal) + 2): [DataTotal] = clsHandlerArray.arrResultsTotal
[DataUniq].ListObject.Resize [DataUniq].ListObject.Range.Resize(UBound(clsHandlerArray.arrResultsSec) + 2): [DataUniq] = clsHandlerArray.arrResultsSec
FILE_ControlPanel.ComplitTimeHandler.Visible = True: FILE_ControlPanel.ComplitTimeHandler.Caption = "Готово!" & " Смет: " & Format(UBound(clsHandlerArray.arrResultsTotal) + 1, "#,##0") _
                                                                                                    & Chr(11) & "Разделов: " & Format(UBound(clsHandlerArray.arrResultsSec), "#,##0") + 1 _
                                                                                                    & Chr(11) & "Строк: " & Format(UBound(clsHandlerArray.arrResultsMain) + 1, "#,##0") _
                                                                                                    & Chr(11) & "Время обработки: " & Format(Timer - Infinity, "0.00 сек")
End Sub
