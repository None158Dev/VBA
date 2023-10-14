VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyGenClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private arr As Variant
Private arr2 As Variant
Private arrPos As Variant
Private arrKeyPos As Variant
Private ArrTotal As Variant
Private DictKey As New Dictionary
Private DictKeysPos As New Dictionary
Private DictKeySum As New Dictionary
Private Sub Class_Initialize()
arr = [DataRes].value2
arr2 = [DataTotal].value2
ReDim arrPos(1 To UBound(arr), 1 To 1)
ReDim arrKeyPos(1 To UBound(arr), 1 To 1)
ReDim ArrTotal(LBound(arr2) To UBound(arr2), 1 To 2)
End Sub
Public Sub KeysPosStart()
Dim i&, CountPos$, G%, CountM&

For i = LBound(arr) To UBound(arr)
    Call KeyPosSm(i, CountPos)
    Call KeyPosUniq(i, CountM)
    DictKeySum(arr(i, 4)) = DictKeySum(arr(i, 4)) + arr(i, 19)
Next i


For i = LBound(arr2) To UBound(arr2)
    ArrTotal(i, 1) = DictKeySum(arr2(i, 2))
    ArrTotal(i, 2) = arr2(i, 4) - DictKeySum(arr2(i, 2))
Next i

End Sub
Private Function KeyPosSm(ByVal index As Long, ByRef CountPos$, Optional G%)
    If Not DictKey.Exists(arr(index, 6)) Then
        DictKey.Add arr(index, 6), ""
        CountPos = 1
    Else
        CountPos = CountPos + 1
    End If
    
    For G = 1 To 4
        If Len(CountPos) = G Then arrPos(index, 1) = arr(index, 6) & Left("/0000", 5 - G) & CountPos
    Next G
End Function
Private Function KeyPosUniq(ByVal index As Long, ByRef CountM&)
Dim rep$
    If Not DictKeysPos.Exists(arr(index, 10)) Then
        DictKeysPos.Add arr(index, 10), ""
        arrKeyPos(index, 1) = arr(index, 10)
    End If
    
    If arr(index, 14) <> "р" Then
        CountM = CountM + 1
        rep = Replace(arr(index, 10), "Р_", "") + CountM
        arrKeyPos(index, 1) = "Д_" & rep
    End If
End Function
Public Property Get arrResultsKey()
arrResultsKey = arrPos
End Property
Public Property Get arrResultsKeyPos()
arrResultsKeyPos = arrKeyPos
End Property
Public Property Get arrResultsTotal()
arrResultsTotal = ArrTotal
End Property
