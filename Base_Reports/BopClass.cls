VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BopClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MainDictKey As New Dictionary
Private DictKeys As New Dictionary
Private DicUniqForKey As New Dictionary
Private DicNotPos As New Dictionary
Private MainArr As Variant
Private ArrRes As Variant
Private Sub Class_Initialize()
MainArr = [pos]
End Sub
Private Sub Class_Terminate()
Set MainDictKey = Nothing
End Sub
Public Sub Start()
Dim i&
Dim ArrTemp As Variant

For i = LBound(MainArr) To UBound(MainArr)
    Call Uniq(i)
    If Not MainDictKey.Exists(MainArr(i, 12)) Then MainDictKey.Add MainArr(i, 12), Array(i, MainDictKey.Count + 1): DictKeys.Add MainArr(i, 12), Array(i, MainDictKey.Count + 1)
Next i

ArrTemp = Array(MainDictKey.Keys, DictKeys.Keys, DictKeys.Items, DicNotPos.Keys, DicNotPos.Items, DicUniqForKey.Items)
ReDim ArrRes(1 To UBound(ArrTemp(0)) + 1, 1 To 15)

For i = LBound(ArrTemp(1)) To UBound(ArrTemp(1))
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 1) = MainArr(ArrTemp(2)(i)(0), 4)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 2) = MainArr(ArrTemp(2)(i)(0), 5)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 3) = MainArr(ArrTemp(2)(i)(0), 6)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 4) = MainArr(ArrTemp(2)(i)(0), 8)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 5) = MainArr(ArrTemp(2)(i)(0), 9)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 6) = MainArr(ArrTemp(2)(i)(0), 10)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 7) = MainArr(ArrTemp(2)(i)(0), 15)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 8) = MainArr(ArrTemp(2)(i)(0), 12)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 9) = MainArr(ArrTemp(2)(i)(0), 17)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 10) = MainArr(ArrTemp(2)(i)(0), 18)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 11) = MainArr(ArrTemp(2)(i)(0), 19)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 12) = MainArr(ArrTemp(2)(i)(0), 22)
    ArrRes(MainDictKey(ArrTemp(1)(i))(1), 13) = MainArr(ArrTemp(2)(i)(0), 20)
Next i

For i = LBound(ArrTemp(3)) To UBound(ArrTemp(3))
    ArrRes(ArrTemp(3)(i), 1) = MainArr(ArrTemp(4)(i)(0), 4)
    ArrRes(ArrTemp(3)(i), 2) = MainArr(ArrTemp(4)(i)(0), 5)
    ArrRes(ArrTemp(3)(i), 3) = MainArr(ArrTemp(4)(i)(0), 6)
    ArrRes(ArrTemp(3)(i), 4) = MainArr(ArrTemp(4)(i)(0), 8)
    ArrRes(ArrTemp(3)(i), 5) = MainArr(ArrTemp(4)(i)(0), 9)
    ArrRes(ArrTemp(3)(i), 6) = MainArr(ArrTemp(4)(i)(0), 10)
    ArrRes(ArrTemp(3)(i), 7) = ArrTemp(4)(i)(1)
    ArrRes(ArrTemp(3)(i), 9) = ArrTemp(4)(i)(2)
Next i

End Sub
Private Function Uniq(index)
If Not DicUniqForKey.Exists(MainArr(index, 4)) Then
    DicUniqForKey.Add MainArr(index, 4), 0
    MainDictKey.Add MainDictKey.Count + 1, ""
    DicNotPos.Add MainDictKey.Count, Array(index, "Смета", MainArr(index, 8))
End If
If Not DicUniqForKey.Exists(MainArr(index, 5)) Then
    DicUniqForKey.Add MainArr(index, 5), 0
    MainDictKey.Add MainDictKey.Count + 1, ""
    DicNotPos.Add MainDictKey.Count, Array(index, "Раздел", MainArr(index, 9))
End If
If Not DicUniqForKey.Exists(MainArr(index, 6)) Then
    DicUniqForKey.Add MainArr(index, 6), 0
    MainDictKey.Add MainDictKey.Count + 1, ""
    DicNotPos.Add MainDictKey.Count, Array(index, "Группа", MainArr(index, 10))
End If

End Function
Public Property Get ItemsRes() As Variant
ItemsRes = ArrRes
End Property




