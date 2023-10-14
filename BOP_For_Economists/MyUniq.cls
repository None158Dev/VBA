VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyUniq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Name class module: MyUniq
Private MyDictKey As New Dictionary
Private arrResP As Variant
Private arrResM As Variant
Private Sub Class_Terminate()
Set MyDictKey = Nothing
End Sub
Public Sub Start(Optional ByVal MainArr As Variant, _
               Optional ByVal OldP As Variant, _
               Optional ByVal OldM As Variant)
Dim i&, G&, CountР&, CountМ&

Dim key As String
Dim KeyGen As String

Dim ArrTemp As Variant
Dim temp As Variant

Dim dicVol As New Dictionary
Dim dicClient As New Dictionary
Dim dicSmUniq As New Dictionary
Dim dicSectionUniq As New Dictionary
Dim dicGroupUniq As New Dictionary
Dim dicOldData As New Dictionary


KeyGen = Fix(Now * 10000000000#)

For i = LBound(MainArr) To UBound(MainArr)
    If MainArr(i, 8) <> "" Then
        key = MainArr(i, 7) & "|" & MainArr(i, 9) & "|" & MainArr(i, 10)
        If Not MyDictKey.Exists(key) Then temp = TypeCount(key, MainArr(i, 7), CountР, CountМ, i)
        dicVol(key) = dicVol(key) + MainArr(i, 11)
        
        If FILE_ControlPanel.MaxUniq.Value = True Then
            If dicClient(key) < MainArr(i, 12) Then dicClient(key) = MainArr(i, 12)
        ElseIf FILE_ControlPanel.MinUniq.Value = True Then
            If Not dicClient.Exists(key) Or dicClient(key) > MainArr(i, 12) And MainArr(i, 12) <> 0 Then dicClient(key) = MainArr(i, 12)
        End If
'====================================================================================================================================================================================
        If Not dicSmUniq.Exists(key) Then dicSmUniq(key) = Chr(10) & MainArr(i, 4) & Chr(10)
        If BlindUniq(dicSmUniq(key), MainArr(i, 4)) = True Then dicSmUniq(key) = dicSmUniq(key) & MainArr(i, 4) & Chr(10)
'====================================================================================================================================================================================
        If Not dicSectionUniq.Exists(key) Then dicSectionUniq(key) = Chr(10) & MainArr(i, 5) & Chr(10)
        If BlindUniq(dicSectionUniq(key), MainArr(i, 5)) = True Then dicSectionUniq(key) = dicSectionUniq(key) & MainArr(i, 5) & Chr(10)
'====================================================================================================================================================================================
        If Not dicGroupUniq.Exists(key) Then dicGroupUniq(key) = Chr(10) & MainArr(i, 6) & Chr(10)
        If BlindUniq(dicGroupUniq(key), MainArr(i, 6)) = True Then dicGroupUniq(key) = dicGroupUniq(key) & MainArr(i, 6) & Chr(10)
    End If
Next i



On Error Resume Next: ReDim arrResP(1 To CountР, 1 To 13)
On Error Resume Next: ReDim arrResM(1 To CountМ, 1 To 13)

For i = LBound(OldP) To UBound(OldP)
     If OldP(i, 10) <> "" Or OldP(i, 11) <> "" Or OldP(i, 12) <> "" Or OldP(i, 13) <> "" Then

        key = OldP(i, 5) & "|" & OldP(i, 6) & "|" & OldP(i, 7)
        If MyDictKey.Exists(key) Then
            arrResP(MyDictKey(key)(0), 10) = OldP(i, 10)
            arrResP(MyDictKey(key)(0), 11) = OldP(i, 11)
            arrResP(MyDictKey(key)(0), 12) = OldP(i, 12)
            arrResP(MyDictKey(key)(0), 13) = OldP(i, 13)
        End If

    End If
Next i

For i = LBound(OldM) To UBound(OldM)
     If OldM(i, 10) <> "" Or OldM(i, 11) <> "" Or OldM(i, 12) <> "" Or OldM(i, 13) <> "" Then

        key = OldM(i, 5) & "|" & OldM(i, 6) & "|" & OldM(i, 7)
        If MyDictKey.Exists(key) Then
            arrResM(MyDictKey(key)(0), 10) = OldM(i, 10)
            arrResM(MyDictKey(key)(0), 11) = OldM(i, 11)
            arrResM(MyDictKey(key)(0), 12) = OldM(i, 12)
            arrResM(MyDictKey(key)(0), 13) = OldM(i, 13)
        End If

    End If
Next i


ArrTemp = Array(MyDictKey.Keys, MyDictKey.Items, dicVol.Items, dicClient.Items, dicSmUniq.Items, dicSectionUniq.Items, dicGroupUniq.Items)


For i = 0 To UBound(ArrTemp(0))
    If ArrTemp(1)(i)(2) Then
        arrResP(ArrTemp(1)(i)(0), 1) = "Р_" & KeyGen + i
        arrResP(ArrTemp(1)(i)(0), 2) = ClearUniq(ArrTemp(4)(i))
        arrResP(ArrTemp(1)(i)(0), 3) = ClearUniq(ArrTemp(5)(i))
        arrResP(ArrTemp(1)(i)(0), 4) = ClearUniq(ArrTemp(6)(i))
        arrResP(ArrTemp(1)(i)(0), 5) = MainArr(ArrTemp(1)(i)(1), 7)
        arrResP(ArrTemp(1)(i)(0), 6) = MainArr(ArrTemp(1)(i)(1), 9)
        arrResP(ArrTemp(1)(i)(0), 7) = MainArr(ArrTemp(1)(i)(1), 10)
        arrResP(ArrTemp(1)(i)(0), 8) = ArrTemp(2)(i)
        arrResP(ArrTemp(1)(i)(0), 9) = ArrTemp(3)(i)
    Else
        arrResM(ArrTemp(1)(i)(0), 1) = "М_" & KeyGen + i
        arrResM(ArrTemp(1)(i)(0), 2) = ClearUniq(ArrTemp(4)(i))
        arrResM(ArrTemp(1)(i)(0), 3) = ClearUniq(ArrTemp(5)(i))
        arrResM(ArrTemp(1)(i)(0), 4) = ClearUniq(ArrTemp(6)(i))
        arrResM(ArrTemp(1)(i)(0), 5) = MainArr(ArrTemp(1)(i)(1), 7)
        arrResM(ArrTemp(1)(i)(0), 6) = MainArr(ArrTemp(1)(i)(1), 9)
        arrResM(ArrTemp(1)(i)(0), 7) = MainArr(ArrTemp(1)(i)(1), 10)
        arrResM(ArrTemp(1)(i)(0), 8) = ArrTemp(2)(i)
        arrResM(ArrTemp(1)(i)(0), 9) = ArrTemp(3)(i)
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
