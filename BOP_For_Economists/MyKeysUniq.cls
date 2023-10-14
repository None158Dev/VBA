VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyKeysUniq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Name class module: MyKeysUniq
Private MyDictKey As New Dictionary
Private MainArr As Variant
Private CoeffArr As Variant
Private ArrKeys As Variant
Private Sub Class_Initialize()
MainArr = [BopSebes].Value2
CoeffArr = [Coefficient].Value2
End Sub
Private Sub Class_Terminate()
Set MyDictKey = Nothing
End Sub
Public Sub Start()
Dim i&
Dim ArrTemp As Variant

For i = LBound(MainArr) To UBound(MainArr)
    If Not MyDictKey.Exists(MainArr(i, 1)) Then MyDictKey.Add MainArr(i, 1), Array(MainArr(i, 4), MyDictKey.Count + 1)
Next i

ArrTemp = Array(MyDictKey.Keys, MyDictKey.Items)

ReDim ArrKeys(1 To MyDictKey.Count, 1 To 5)

For i = LBound(CoeffArr) To UBound(CoeffArr)
    If MyDictKey.Exists(CoeffArr(i, 1)) Then
        ArrKeys(MyDictKey(CoeffArr(i, 1))(1), 3) = CoeffArr(i, 3)
        ArrKeys(MyDictKey(CoeffArr(i, 1))(1), 4) = CoeffArr(i, 4)
        ArrKeys(MyDictKey(CoeffArr(i, 1))(1), 5) = CoeffArr(i, 5)
    End If
Next i

For i = 1 To MyDictKey.Count
    ArrKeys(i, 1) = ArrTemp(0)(i - 1)
    ArrKeys(i, 2) = ArrTemp(1)(i - 1)(0)
Next i
End Sub
Public Property Get UniqKeys() As Variant
UniqKeys = ArrKeys
End Property
