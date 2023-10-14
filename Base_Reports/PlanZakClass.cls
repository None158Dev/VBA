VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanZakClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MainDictKey As New Dictionary
Private DicSection As New Dictionary
Private DicUniqSm As New Dictionary
Private DictKeys As New Dictionary

Private MainArr As Variant
Private PercentArr As Variant
Private ArrRes As Variant
Private Sub Class_Initialize()
MainArr = [WO]
PercentArr = [ForPlanZak]
End Sub
Private Sub Class_Terminate()
Set MainDictKey = Nothing
End Sub
Public Sub Start()
Dim i&
Dim ArrTemp As Variant
Dim arrTempVal As Variant
Dim arrTempResVal As Variant

Dim key As String
Dim PosSmeta As Long
Dim DicVol As New Dictionary
Dim DicClient As New Dictionary
Dim DicMarket As New Dictionary

Call KeysSection

For i = LBound(MainArr) To UBound(MainArr)
    If MainArr(i, 18) <> "р" And MainArr(i, 62) = "True" And DicSection.Exists(MainArr(i, 24)) Then
        key = MainArr(i, 26) & "|" & MainArr(i, 27) & "|" & MainArr(i, 18) & "|" & MainArr(i, 20) & "|" & MainArr(i, 21)
        If Not DictKeys.Exists(key) Then DictKeys.Add key, Array(i, DictKeys.Count)
        
        DicVol(key) = DicVol(key) + MainArr(i, 7)
        DicClient(key) = DicClient(key) + MainArr(i, 15)
        DicMarket(key) = DicMarket(key) + MainArr(i, 16)
        
    End If

Next i

ArrTemp = Array(DictKeys.Keys, DictKeys.Items)
arrTempVal = Array(DicVol.Items, DicClient.Items, DicMarket.Items)

ReDim arrTempResVal(0 To UBound(arrTempVal(0)), 1 To 5)


For i = LBound(arrTempVal(0)) To UBound(arrTempVal(0))

    arrTempResVal(i, 1) = arrTempVal(0)(i) * DicSection(MainArr(ArrTemp(1)(i)(0), 24))(1)
    arrTempResVal(i, 2) = arrTempVal(0)(i) * DicSection(MainArr(ArrTemp(1)(i)(0), 24))(2)
    arrTempResVal(i, 3) = arrTempVal(0)(i) * DicSection(MainArr(ArrTemp(1)(i)(0), 24))(3)
    arrTempResVal(i, 4) = arrTempVal(1)(i) / arrTempVal(0)(i)
    arrTempResVal(i, 5) = arrTempVal(2)(i) / arrTempVal(0)(i)
    If MainArr(ArrTemp(1)(i)(0), 21) Like "компл*" Or MainArr(ArrTemp(1)(i)(0), 21) Like "шт*" Then
        arrTempResVal(i, 1) = Round(arrTempResVal(i, 1), 0)
        arrTempResVal(i, 2) = Round(arrTempResVal(i, 2), 0)
        arrTempResVal(i, 3) = Round(arrTempResVal(i, 3), 0)
        
        If arrTempResVal(i, 1) = 0 And arrTempResVal(i, 2) = 0 And arrTempResVal(i, 3) = 0 Then
            DictKeys.Remove (ArrTemp(0)(i))
        End If

    End If
    
    If DictKeys.Exists(ArrTemp(0)(i)) Then
        If Not MainDictKey.Exists(ArrTemp(0)(i)) Then
            Call UniqSm(ArrTemp(1)(i)(0), PosSmeta)
            MainDictKey.Add ArrTemp(0)(i), Array(MainDictKey.Count + 1, ArrTemp(1)(i)(0), PosSmeta)
        End If
    End If
    
Next i

ArrTemp = Array(DictKeys.Keys, DictKeys.Items, DicUniqSm.Keys, DicUniqSm.Items)

ReDim ArrRes(1 To MainDictKey.Count, 1 To 22)

For i = LBound(ArrTemp(0)) To UBound(ArrTemp(0))
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 1) = MainDictKey(ArrTemp(0)(i))(2)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 2) = MainArr(ArrTemp(1)(i)(0), 23)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 3) = MainArr(ArrTemp(1)(i)(0), 24)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 4) = MainArr(ArrTemp(1)(i)(0), 26)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 5) = MainArr(ArrTemp(1)(i)(0), 27)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 6) = MainArr(ArrTemp(1)(i)(0), 18)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 7) = MainArr(ArrTemp(1)(i)(0), 20)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 8) = MainArr(ArrTemp(1)(i)(0), 21)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 9) = arrTempResVal(ArrTemp(1)(i)(1), 4)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 10) = arrTempResVal(ArrTemp(1)(i)(1), 5)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 11) = arrTempResVal(ArrTemp(1)(i)(1), 1)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 12) = ArrRes(MainDictKey(ArrTemp(0)(i))(0), 9) * ArrRes(MainDictKey(ArrTemp(0)(i))(0), 11)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 13) = ArrRes(MainDictKey(ArrTemp(0)(i))(0), 10) * ArrRes(MainDictKey(ArrTemp(0)(i))(0), 11)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 14) = ArrRes(MainDictKey(ArrTemp(0)(i))(0), 10) * ArrRes(MainDictKey(ArrTemp(0)(i))(0), 11)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 15) = arrTempResVal(ArrTemp(1)(i)(1), 2)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 16) = ArrRes(MainDictKey(ArrTemp(0)(i))(0), 9) * ArrRes(MainDictKey(ArrTemp(0)(i))(0), 15)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 17) = ArrRes(MainDictKey(ArrTemp(0)(i))(0), 10) * ArrRes(MainDictKey(ArrTemp(0)(i))(0), 15)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 18) = ArrRes(MainDictKey(ArrTemp(0)(i))(0), 10) * ArrRes(MainDictKey(ArrTemp(0)(i))(0), 15)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 19) = arrTempResVal(ArrTemp(1)(i)(1), 3)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 20) = ArrRes(MainDictKey(ArrTemp(0)(i))(0), 9) * ArrRes(MainDictKey(ArrTemp(0)(i))(0), 19)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 21) = ArrRes(MainDictKey(ArrTemp(0)(i))(0), 10) * ArrRes(MainDictKey(ArrTemp(0)(i))(0), 19)
    ArrRes(MainDictKey(ArrTemp(0)(i))(0), 22) = ArrRes(MainDictKey(ArrTemp(0)(i))(0), 10) * ArrRes(MainDictKey(ArrTemp(0)(i))(0), 19)
Next i

For i = LBound(ArrTemp(2)) To UBound(ArrTemp(2))
    ArrRes(MainDictKey(ArrTemp(2)(i)), 1) = ArrTemp(3)(i)(2)
    ArrRes(MainDictKey(ArrTemp(2)(i)), 2) = MainArr(ArrTemp(3)(i)(0), 23)
    ArrRes(MainDictKey(ArrTemp(2)(i)), 3) = MainArr(ArrTemp(3)(i)(0), 24)
    ArrRes(MainDictKey(ArrTemp(2)(i)), 4) = MainArr(ArrTemp(3)(i)(0), 26)
    ArrRes(MainDictKey(ArrTemp(2)(i)), 6) = ArrTemp(3)(i)(1)
Next i

End Sub
Private Function UniqSm(ByVal index As Long, ByRef PosSmeta As Long)
If Not DicUniqSm.Exists(MainArr(index, 23)) Then
    PosSmeta = DicUniqSm.Count + 1
    DicUniqSm.Add MainArr(index, 23), Array(index, "Смета", PosSmeta)
    MainDictKey.Add MainArr(index, 23), MainDictKey.Count + 1
End If
End Function
Private Function KeysSection(Optional i As Long)
For i = LBound(PercentArr) To UBound(PercentArr)
    If Not DicSection.Exists(PercentArr(i, 1)) Then
        DicSection.Add PercentArr(i, 1), Array(DicSection.Count + 1, PercentArr(i, 2), PercentArr(i, 3), PercentArr(i, 4))
    End If
Next i
End Function
Public Property Get ItemsRes() As Variant
ItemsRes = ArrRes
End Property

