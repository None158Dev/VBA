VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Name class module: MyCount
Private MainArr As Variant
Private OldP As Variant
Private OldM As Variant
Private AllSmCountComp As Variant
Private ArrCountPM As Variant
Private ArrCountSM As Variant
Private Sub Class_Initialize()
MainArr = [BopSebes].Value2
OldP = [pos_UniqP].Value2
OldM = [pos_UniqM].Value2
AllSmCountComp = [CountSm].Value2
End Sub
Public Sub CountCompAll()
Dim i&, CountCompliteP&, CountClientP&, CountCompliteM&, CountClientM&

ReDim ArrCountPM(1 To 2, 1 To 5)


For i = LBound(OldP) To UBound(OldP)
    If OldP(i, 14) = "Обновлено" Then CountCompliteP = CountCompliteP + 1
    If OldP(i, 14) = "Цена по заказчику" Then CountClientP = CountClientP + 1
Next i

For i = LBound(OldM) To UBound(OldM)
    If OldM(i, 14) = "Обновлено" Then CountCompliteM = CountCompliteM + 1
    If OldM(i, 14) = "Цена по заказчику" Then CountClientM = CountClientM + 1
Next i

ArrCountPM(1, 1) = UBound(OldP)
ArrCountPM(2, 1) = UBound(OldM)
ArrCountPM(1, 2) = CountCompliteP
ArrCountPM(2, 2) = CountCompliteM
ArrCountPM(1, 3) = UBound(OldP) - CountCompliteP
ArrCountPM(2, 3) = UBound(OldM) - CountCompliteM
ArrCountPM(1, 4) = CountCompliteP / UBound(OldP)
ArrCountPM(2, 4) = CountCompliteM / UBound(OldM)
ArrCountPM(1, 5) = CountClientP
ArrCountPM(2, 5) = CountClientM

End Sub
Public Sub CountCompSM()
Dim i&
Dim key As String
Dim ArrTemp As Variant
Dim MyDictKey As New Dictionary

ReDim ArrCountSM(1 To UBound(AllSmCountComp), 1 To 9)

For i = LBound(MainArr) To UBound(MainArr)
    If MainArr(i, 8) <> "" Then
        key = "|" & MainArr(i, 1) & "|" & MainArr(i, 7) & "|" & MainArr(i, 9) & "|" & MainArr(i, 10)
        If Not MyDictKey.Exists(key) Then MyDictKey.Add key, "|" & MainArr(i, 1) & "|" & MainArr(i, 19) & "|" & Chr(10) & "|" & MainArr(i, 1) & "|" & MainArr(i, 7) & "|" & MainArr(i, 19) & "|"
    End If
Next i

ArrTemp = MyDictKey.Items

For i = LBound(AllSmCountComp) To UBound(AllSmCountComp)
    ArrCountSM(i, 1) = UBound(Filter(MyDictKey.Items, "|" & AllSmCountComp(i, 1) & "|", True)) + 1
    ArrCountSM(i, 2) = ArrCountSM(i, 1) - (UBound(Filter(MyDictKey.Items, "|" & AllSmCountComp(i, 1) & "|Обновлено|", True)) + 1)
    ArrCountSM(i, 4) = UBound(Filter(MyDictKey.Keys, "|" & AllSmCountComp(i, 1) & "|р|", True)) + 1
    ArrCountSM(i, 5) = UBound(Filter(MyDictKey.Items, Chr(10) & "|" & AllSmCountComp(i, 1) & "|р|" & "Не обновлено|", True)) + 1
    ArrCountSM(i, 7) = UBound(Filter(MyDictKey.Keys, "|" & AllSmCountComp(i, 1) & "|м|", True)) + 1
    ArrCountSM(i, 8) = UBound(Filter(MyDictKey.Items, Chr(10) & "|" & AllSmCountComp(i, 1) & "|м|" & "Не обновлено|", True)) + 1
    If ArrCountSM(i, 1) = 0 Then ArrCountSM(i, 3) = 1 Else ArrCountSM(i, 3) = (ArrCountSM(i, 1) - ArrCountSM(i, 2)) / ArrCountSM(i, 1)
    If ArrCountSM(i, 4) = 0 Then ArrCountSM(i, 6) = 1 Else ArrCountSM(i, 6) = (ArrCountSM(i, 4) - ArrCountSM(i, 5)) / ArrCountSM(i, 4)
    If ArrCountSM(i, 7) = 0 Then ArrCountSM(i, 9) = 1 Else ArrCountSM(i, 9) = (ArrCountSM(i, 7) - ArrCountSM(i, 8)) / ArrCountSM(i, 7)
Next i

End Sub
Public Property Get UniqCompPM() As Variant
UniqCompPM = ArrCountPM
End Property
Public Property Get UniqCompSM() As Variant
UniqCompSM = ArrCountSM
End Property

