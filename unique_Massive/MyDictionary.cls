VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MyDictKey As New Dictionary
Private arrResP As Variant
Private arrResM As Variant
Private Sub Class_Terminate()
Set MyDictKey = Nothing
End Sub
Public Sub add(Optional ByVal MainArr As Variant)
Dim i&, G&, CountР&, CountМ&

Dim key As String
Dim KeyGen As String

Dim arrTemp As Variant
Dim Temp As Variant

Dim dicVol As New Dictionary
Dim dicClient As New Dictionary
Dim dicMarket As New Dictionary
Dim dicSmUniq As New Dictionary
Dim dicSectionUniq As New Dictionary
Dim dicGroupUniq As New Dictionary

KeyGen = Fix(Now * 10000000000#)

For i = LBound(MainArr) To UBound(MainArr)
    key = MainArr(i, 4) & "|" & MainArr(i, 5) & "|" & MainArr(i, 6)
    If Not MyDictKey.Exists(key) Then Temp = TypeCount(key, MainArr(i, 4), CountР, CountМ, i)
    dicVol(key) = dicVol(key) + MainArr(i, 7)
    
    If FILE_ControlPanel.MaxUniq.Value = True Then
        If dicClient(key) < MainArr(i, 8) Then dicClient(key) = MainArr(i, 8)
        If dicMarket(key) < MainArr(i, 9) Then dicMarket(key) = MainArr(i, 9)
    ElseIf FILE_ControlPanel.MinUniq.Value = True Then
        If Not dicClient.Exists(key) Or dicClient(key) > MainArr(i, 8) And MainArr(i, 8) <> 0 Then dicClient(key) = MainArr(i, 8)
        If Not dicMarket.Exists(key) Or dicMarket(key) > MainArr(i, 9) And MainArr(i, 9) <> 0 Then dicMarket(key) = MainArr(i, 9)
    End If
'====================================================================================================================================================================================
    If Not dicSmUniq.Exists(key) Then dicSmUniq(key) = Chr(10) & MainArr(i, 1) & Chr(10)
    If BlindUniq(dicSmUniq(key), MainArr(i, 1)) = True Then dicSmUniq(key) = dicSmUniq(key) & MainArr(i, 1) & Chr(10)
'====================================================================================================================================================================================
    If Not dicSectionUniq.Exists(key) Then dicSectionUniq(key) = Chr(10) & MainArr(i, 2) & Chr(10)
    If BlindUniq(dicSectionUniq(key), MainArr(i, 2)) = True Then dicSectionUniq(key) = dicSectionUniq(key) & MainArr(i, 2) & Chr(10)
'====================================================================================================================================================================================
    If Not dicGroupUniq.Exists(key) Then dicGroupUniq(key) = Chr(10) & MainArr(i, 3) & Chr(10)
    If BlindUniq(dicGroupUniq(key), MainArr(i, 3)) = True Then dicGroupUniq(key) = dicGroupUniq(key) & MainArr(i, 3) & Chr(10)

Next i

arrTemp = Array(MyDictKey.Keys, MyDictKey.Items, dicVol.Items, dicClient.Items, dicMarket.Items, dicSmUniq.Items, dicSectionUniq.Items, dicGroupUniq.Items)

On Error Resume Next: ReDim arrResP(1 To CountР, 1 To 10)
On Error Resume Next: ReDim arrResM(1 To CountМ, 1 To 10)

For i = 0 To UBound(arrTemp(0))
    If arrTemp(1)(i)(2) Then
        arrResP(arrTemp(1)(i)(0), 1) = "Р_" & KeyGen + i
        arrResP(arrTemp(1)(i)(0), 2) = ClearUniq(arrTemp(5)(i))
        arrResP(arrTemp(1)(i)(0), 3) = ClearUniq(arrTemp(6)(i))
        arrResP(arrTemp(1)(i)(0), 4) = ClearUniq(arrTemp(7)(i))
        arrResP(arrTemp(1)(i)(0), 5) = MainArr(arrTemp(1)(i)(1), 4)
        arrResP(arrTemp(1)(i)(0), 6) = MainArr(arrTemp(1)(i)(1), 5)
        arrResP(arrTemp(1)(i)(0), 7) = MainArr(arrTemp(1)(i)(1), 6)
        arrResP(arrTemp(1)(i)(0), 8) = arrTemp(2)(i)
        arrResP(arrTemp(1)(i)(0), 9) = arrTemp(3)(i)
        arrResP(arrTemp(1)(i)(0), 10) = arrTemp(4)(i)
    Else
        arrResM(arrTemp(1)(i)(0), 1) = "М_" & KeyGen + i
        arrResM(arrTemp(1)(i)(0), 2) = ClearUniq(arrTemp(5)(i))
        arrResM(arrTemp(1)(i)(0), 3) = ClearUniq(arrTemp(6)(i))
        arrResM(arrTemp(1)(i)(0), 4) = ClearUniq(arrTemp(7)(i))
        arrResM(arrTemp(1)(i)(0), 5) = MainArr(arrTemp(1)(i)(1), 4)
        arrResM(arrTemp(1)(i)(0), 6) = MainArr(arrTemp(1)(i)(1), 5)
        arrResM(arrTemp(1)(i)(0), 7) = MainArr(arrTemp(1)(i)(1), 6)
        arrResM(arrTemp(1)(i)(0), 8) = arrTemp(2)(i)
        arrResM(arrTemp(1)(i)(0), 9) = arrTemp(3)(i)
        arrResM(arrTemp(1)(i)(0), 10) = arrTemp(4)(i)
    End If
Next i

End Sub
Private Function TypeCount(ByVal key As String, ByVal TypePos As String, ByRef CountР As Long, ByRef CountМ As Long, indexMain As Long)
If TypePos = "р" Then CountР = CountР + 1: MyDictKey.add key, Array(CountР, indexMain, True)
If TypePos <> "р" Then CountМ = CountМ + 1: MyDictKey.add key, Array(CountМ, indexMain, False)
End Function
Private Function BlindUniq(ByVal key As String, ByVal item As String) As Boolean: BlindUniq = False
If Not key Like "*" & Chr(10) & item & Chr(10) & "*" Then BlindUniq = True
End Function
Private Function ClearUniq(ByVal Value As String) As String
ClearUniq = Mid(Value, 2, Len(Value) - 2)
End Function
Public Property Get ItemsP() As Variant
ItemsP = arrResP
End Property
Public Property Get ItemsM() As Variant
ItemsM = arrResM
End Property
