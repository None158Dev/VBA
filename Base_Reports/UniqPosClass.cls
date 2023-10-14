VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UniqPosClass"
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
Public Sub Start(Optional ByVal MainArr As Variant)
Dim i&, G&, CountР&, CountМ&

Dim key As String
Dim KeyGen As String

Dim ArrTemp As Variant
Dim Temp As Variant

Dim DicVol As New Dictionary
Dim DicClient As New Dictionary
Dim DicMarket As New Dictionary
Dim dicSmUniq As New Dictionary
Dim dicSectionUniq As New Dictionary
Dim dicGroupUniq As New Dictionary

KeyGen = Fix(Now * 10000000000#)

For i = LBound(MainArr) To UBound(MainArr)
    key = MainArr(i, 15) & "|" & MainArr(i, 17) & "|" & MainArr(i, 18)
    If Not MyDictKey.Exists(key) Then Temp = TypeCount(key, MainArr(i, 15), CountР, CountМ, i)
    DicVol(key) = DicVol(key) + MainArr(i, 19)
    
    If FILE_FRM_Search.MaxUniq.Value = True Then
        If DicClient(key) < MainArr(i, 22) Then DicClient(key) = MainArr(i, 22)
        If DicMarket(key) < MainArr(i, 21) Then DicMarket(key) = MainArr(i, 21)
    ElseIf FILE_FRM_Search.MinUniq.Value = True Then
        If Not DicClient.Exists(key) Or DicClient(key) > MainArr(i, 22) And MainArr(i, 22) <> 0 Then DicClient(key) = MainArr(i, 22)
        If Not DicMarket.Exists(key) Or DicMarket(key) > MainArr(i, 21) And MainArr(i, 21) <> 0 Then DicMarket(key) = MainArr(i, 21)
    End If
'====================================================================================================================================================================================
If Not dicSmUniq.Exists(key) Then dicSmUniq(key) = Chr(10) & MainArr(i, 8) & Chr(10)
If BlindUniq(dicSmUniq(key), MainArr(i, 8)) = True Then dicSmUniq(key) = dicSmUniq(key) & MainArr(i, 8) & Chr(10)
'====================================================================================================================================================================================
If Not dicSectionUniq.Exists(key) Then dicSectionUniq(key) = Chr(10) & MainArr(i, 9) & Chr(10)
If BlindUniq(dicSectionUniq(key), MainArr(i, 9)) = True Then dicSectionUniq(key) = dicSectionUniq(key) & MainArr(i, 9) & Chr(10)
'====================================================================================================================================================================================
If Not dicGroupUniq.Exists(key) Then dicGroupUniq(key) = Chr(10) & MainArr(i, 10) & Chr(10)
If BlindUniq(dicGroupUniq(key), MainArr(i, 10)) = True Then dicGroupUniq(key) = dicGroupUniq(key) & MainArr(i, 10) & Chr(10)

Next i

ArrTemp = Array(MyDictKey.Keys, MyDictKey.Items, DicVol.Items, DicClient.Items, DicMarket.Items, dicSmUniq.Items, dicSectionUniq.Items, dicGroupUniq.Items)

On Error Resume Next: ReDim arrResP(1 To CountР, 1 To 10)
On Error Resume Next: ReDim arrResM(1 To CountМ, 1 To 10)

For i = 0 To UBound(ArrTemp(0))
    If ArrTemp(1)(i)(2) Then
        arrResP(ArrTemp(1)(i)(0), 1) = "Р_" & KeyGen + i
        arrResP(ArrTemp(1)(i)(0), 2) = ClearUniq(ArrTemp(5)(i))
        arrResP(ArrTemp(1)(i)(0), 3) = ClearUniq(ArrTemp(6)(i))
        arrResP(ArrTemp(1)(i)(0), 4) = ClearUniq(ArrTemp(7)(i))
        arrResP(ArrTemp(1)(i)(0), 5) = MainArr(ArrTemp(1)(i)(1), 15)
        arrResP(ArrTemp(1)(i)(0), 6) = MainArr(ArrTemp(1)(i)(1), 17)
        arrResP(ArrTemp(1)(i)(0), 7) = MainArr(ArrTemp(1)(i)(1), 18)
        arrResP(ArrTemp(1)(i)(0), 8) = ArrTemp(2)(i)
        arrResP(ArrTemp(1)(i)(0), 9) = ArrTemp(3)(i)
        arrResP(ArrTemp(1)(i)(0), 10) = ArrTemp(4)(i)
    Else
        arrResM(ArrTemp(1)(i)(0), 1) = "М_" & KeyGen + i
        arrResM(ArrTemp(1)(i)(0), 2) = ClearUniq(ArrTemp(5)(i))
        arrResM(ArrTemp(1)(i)(0), 3) = ClearUniq(ArrTemp(6)(i))
        arrResM(ArrTemp(1)(i)(0), 4) = ClearUniq(ArrTemp(7)(i))
        arrResM(ArrTemp(1)(i)(0), 5) = MainArr(ArrTemp(1)(i)(1), 15)
        arrResM(ArrTemp(1)(i)(0), 6) = MainArr(ArrTemp(1)(i)(1), 17)
        arrResM(ArrTemp(1)(i)(0), 7) = MainArr(ArrTemp(1)(i)(1), 18)
        arrResM(ArrTemp(1)(i)(0), 8) = ArrTemp(2)(i)
        arrResM(ArrTemp(1)(i)(0), 9) = ArrTemp(3)(i)
        arrResM(ArrTemp(1)(i)(0), 10) = ArrTemp(4)(i)
    End If
Next i

End Sub
Private Function TypeCount(ByVal key As String, ByVal TypePos As String, ByRef CountР As Long, ByRef CountМ As Long, indexMain As Long)
If TypePos = "р" Then CountР = CountР + 1: MyDictKey.Add key, Array(CountР, indexMain, True)
If TypePos <> "р" Then CountМ = CountМ + 1: MyDictKey.Add key, Array(CountМ, indexMain, False)
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
