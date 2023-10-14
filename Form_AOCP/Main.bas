Attribute VB_Name = "Main"
Option Explicit
Public sh_Act As Worksheet
Public sh_Main As Worksheet
Public ArrTable_IndexAct() As Variant
Public dic_UniqAct As New Dictionary
Public ArrFilter As Variant
Private Sub ObjectUpdate()
Set sh_Act = [_NumberActB].Worksheet
Set sh_Main = [_SearchSheet].Worksheet
End Sub
Private Sub PDFUnload()
Dim i&
Dim ArrTable_Main() As Variant
'================================================
Application.ScreenUpdating = False
ObjectUpdate
ArrTable_Main = sh_Main.Range("A7:B" & sh_Main.Cells(Rows.Count, 2).End(xlUp).Row).Value2
'===========================================================================
    For i = 1 To UBound(ArrTable_Main)
        dic_UniqAct.Item(ArrTable_Main(i, 1) & " | " & ArrTable_Main(i, 2)) = dic_UniqAct.Item(ArrTable_Main(i, 1) & " | " & ArrTable_Main(i, 2))
    Next i
'===========================================================================
Application.ScreenUpdating = True
End Sub
Private Sub Path()
Dim inpPath As String
inpPath = InputBox("Введите путь в формате:" & Chr(13) & "Z:\ПТО\0) Акты pdf" & Chr(13) & "И сохраните файл!")
'===========================================================================
If inpPath = "" Then: Exit Sub
[_Path].Value2 = inpPath
End Sub
Private Sub UserShow()
Application.Calculation = xlAutomatic
Search_Form.Show
End Sub
Sub Регсчетчика11_Изменение()
Dim i&
Dim ArrObject As Variant
Dim Horiz As Double
'===========================================================================
Application.ScreenUpdating = False
ArrObject = Array([_HorizOne], [_HorizTwo], [_HorizThree], [_HorizFour], [_HorizSix], [_HorizSeven], [_HorizApp])
'===========================================================================
For i = 0 To UBound(ArrObject)
    ArrObject(i).UnMerge
    ArrObject(i).HorizontalAlignment = 7
    ArrObject(i).EntireRow.AutoFit
    Horiz = ArrObject(i).RowHeight
    ArrObject(i).Merge
    ArrObject(i).RowHeight = Horiz
    ArrObject(i).HorizontalAlignment = xlLeft
    ArrObject(i).VerticalAlignment = xlCenter
Next i
Application.ScreenUpdating = True
End Sub
