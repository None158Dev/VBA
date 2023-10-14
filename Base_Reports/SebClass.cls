VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SebClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MainDictKey As New Dictionary
Private DictLvl As New Dictionary
Private DictKeyPodr As New Dictionary
Private DictKeys As New Dictionary
Private DictKeyCoeff As New Dictionary
Private MainArr As Variant
Private KeyPodrArr As Variant
Private ArrTemp As Variant
Private ArrRes As Variant
Private SebesCoeffArr As Variant
Private Sub Class_Initialize()
MainArr = [WO]
KeyPodrArr = [Cntr]
SebesCoeffArr = [SebesCoeff]
Call KeyPodr
Call KeyCoeff
End Sub
Private Sub Class_Terminate()
Set MainDictKey = Nothing
End Sub
Public Sub Start()
Dim i&
Dim ArrTemp As Variant
Dim key As String

For i = LBound(MainArr) To UBound(MainArr)
    key = MainArr(i, 4) & "|" & MainArr(i, 3)
    If Not MainDictKey.Exists(key) And DictKeyCoeff.Exists(MainArr(i, 24)) Then
        Call KeyLvl(i)
        MainDictKey.Add key, ""
        DictKeys.Add key, Array(i, MainDictKey.Count, SebesCoeffArr(DictKeyCoeff(MainArr(i, 24)), 2))
    End If
    
Next i

ArrTemp = Array(MainDictKey.Keys, DictKeys.Items, DictLvl.Items)

ReDim ArrRes(1 To UBound(ArrTemp(0)) + 1, 1 To 23)

For i = LBound(ArrTemp(1)) To UBound(ArrTemp(1))

    ArrRes(ArrTemp(1)(i)(1), 1) = MainArr(ArrTemp(1)(i)(0), 23)
    ArrRes(ArrTemp(1)(i)(1), 2) = MainArr(ArrTemp(1)(i)(0), 24)
    ArrRes(ArrTemp(1)(i)(1), 3) = MainArr(ArrTemp(1)(i)(0), 25)
    ArrRes(ArrTemp(1)(i)(1), 4) = MainArr(ArrTemp(1)(i)(0), 26)
    ArrRes(ArrTemp(1)(i)(1), 5) = MainArr(ArrTemp(1)(i)(0), 27)
    ArrRes(ArrTemp(1)(i)(1), 6) = MainArr(ArrTemp(1)(i)(0), 28)
    ArrRes(ArrTemp(1)(i)(1), 7) = MainArr(ArrTemp(1)(i)(0), 3)
    ArrRes(ArrTemp(1)(i)(1), 8) = KeyPodrArr(DictKeyPodr(MainArr(ArrTemp(1)(i)(0), 3)), 4)
    ArrRes(ArrTemp(1)(i)(1), 9) = MainArr(ArrTemp(1)(i)(0), 18)
    ArrRes(ArrTemp(1)(i)(1), 10) = MainArr(ArrTemp(1)(i)(0), 5)
    ArrRes(ArrTemp(1)(i)(1), 11) = MainArr(ArrTemp(1)(i)(0), 4)
    ArrRes(ArrTemp(1)(i)(1), 12) = MainArr(ArrTemp(1)(i)(0), 20)
    ArrRes(ArrTemp(1)(i)(1), 13) = MainArr(ArrTemp(1)(i)(0), 21)
    ArrRes(ArrTemp(1)(i)(1), 14) = MainArr(ArrTemp(1)(i)(0), 7)
    ArrRes(ArrTemp(1)(i)(1), 15) = MainArr(ArrTemp(1)(i)(0), 7) * ArrTemp(1)(i)(2)
    ArrRes(ArrTemp(1)(i)(1), 16) = MainArr(ArrTemp(1)(i)(0), 13)
    ArrRes(ArrTemp(1)(i)(1), 17) = MainArr(ArrTemp(1)(i)(0), 15)
    ArrRes(ArrTemp(1)(i)(1), 18) = ArrRes(ArrTemp(1)(i)(1), 15) * MainArr(ArrTemp(1)(i)(0), 13)
    ArrRes(ArrTemp(1)(i)(1), 19) = MainArr(ArrTemp(1)(i)(0), 12)
    ArrRes(ArrTemp(1)(i)(1), 20) = MainArr(ArrTemp(1)(i)(0), 6)
    ArrRes(ArrTemp(1)(i)(1), 21) = MainArr(ArrTemp(1)(i)(0), 14)
    ArrRes(ArrTemp(1)(i)(1), 22) = MainArr(ArrTemp(1)(i)(0), 16)
    ArrRes(ArrTemp(1)(i)(1), 23) = ArrRes(ArrTemp(1)(i)(1), 15) * MainArr(ArrTemp(1)(i)(0), 14)
    
Next i

For i = LBound(ArrTemp(2)) To UBound(ArrTemp(2))
    ArrRes(ArrTemp(2)(i)(0), 1) = MainArr(ArrTemp(2)(i)(1), 23)
    ArrRes(ArrTemp(2)(i)(0), 2) = MainArr(ArrTemp(2)(i)(1), 24)
    ArrRes(ArrTemp(2)(i)(0), 3) = MainArr(ArrTemp(2)(i)(1), 25)
    ArrRes(ArrTemp(2)(i)(0), 4) = MainArr(ArrTemp(2)(i)(1), 26)
    ArrRes(ArrTemp(2)(i)(0), 5) = MainArr(ArrTemp(2)(i)(1), 27)
    ArrRes(ArrTemp(2)(i)(0), 6) = MainArr(ArrTemp(2)(i)(1), 28)
    ArrRes(ArrTemp(2)(i)(0), 9) = ArrTemp(2)(i)(2)
    ArrRes(ArrTemp(2)(i)(0), 12) = ArrTemp(2)(i)(3)
Next i

End Sub
Private Function KeyLvl(index)

If Not DictLvl.Exists(MainArr(index, 23)) Then
DictLvl.Add MainArr(index, 23), Array(MainDictKey.Count + 1, index, "Смета", MainArr(index, 26))
MainDictKey.Add MainDictKey.Count + 1, Array(index, MainDictKey.Count + 1)
End If

If Not DictLvl.Exists(MainArr(index, 24)) Then
DictLvl.Add MainArr(index, 24), Array(MainDictKey.Count + 1, index, "Раздел", MainArr(index, 27))
MainDictKey.Add MainDictKey.Count + 1, Array(index, MainDictKey.Count + 1)
End If

If Not DictLvl.Exists(MainArr(index, 25)) Then
DictLvl.Add MainArr(index, 25), Array(MainDictKey.Count + 1, index, "Группа", MainArr(index, 28))
MainDictKey.Add MainDictKey.Count + 1, Array(index, MainDictKey.Count + 1)
End If

End Function
Private Function KeyPodr(Optional i As Long)
For i = LBound(KeyPodrArr) To UBound(KeyPodrArr)
    DictKeyPodr.Add KeyPodrArr(i, 3), i
Next i
End Function
Private Function KeyCoeff(Optional i As Long)
For i = LBound(SebesCoeffArr) To UBound(SebesCoeffArr)
    DictKeyCoeff.Add SebesCoeffArr(i, 1), i
Next i
End Function
Public Property Get ItemsRes() As Variant
ItemsRes = ArrRes
End Property
