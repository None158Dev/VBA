VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Name class module: MySearch
Private MyDictKey As New Dictionary
Private MainArr As Variant
Private OldP As Variant
Private OldM As Variant
Private ArrMarket As Variant
Private ArrStatus As Variant
Private Sub Class_Initialize()
MainArr = [BopSebes].Value2
OldP = [pos_UniqP].Value2
OldM = [pos_UniqM].Value2
ReDim ArrMarket(1 To UBound(MainArr), 1 To 1)
ReDim ArrStatus(1 To UBound(MainArr), 1 To 1)
End Sub
Private Sub Class_Terminate()
Set MyDictKey = Nothing
End Sub
Public Sub Start()
Dim i&
Dim key As String

For i = LBound(OldP) To UBound(OldP)
    key = OldP(i, 5) & "|" & OldP(i, 6) & "|" & OldP(i, 7)
    If OldP(i, 14) = "Цена по заказчику" Then MyDictKey.Add key, Array(OldP(i, 10), OldP(i, 14))
    If OldP(i, 14) = "Обновлено" Then MyDictKey.Add key, Array(OldP(i, 12), OldP(i, 14))
Next i


For i = LBound(OldM) To UBound(OldM)
    key = OldM(i, 5) & "|" & OldM(i, 6) & "|" & OldM(i, 7)
    If OldM(i, 14) = "Цена по заказчику" Then MyDictKey.Add key, Array(OldM(i, 10), OldM(i, 14))
    If OldM(i, 14) = "Обновлено" Then MyDictKey.Add key, Array(OldM(i, 12), OldM(i, 14))
Next i



For i = LBound(MainArr) To UBound(MainArr)
    If MainArr(i, 8) <> "" Then
        key = MainArr(i, 7) & "|" & MainArr(i, 9) & "|" & MainArr(i, 10)
        If Not MyDictKey.Exists(key) Then
            ArrStatus(i, 1) = "Не обновлено"
        Else
            ArrMarket(i, 1) = MyDictKey(key)(0): ArrStatus(i, 1) = MyDictKey(key)(1)
        End If
    End If
Next i

End Sub
Public Property Get Market() As Variant
Market = ArrMarket
End Property
Public Property Get Status() As Variant
Status = ArrStatus
End Property
