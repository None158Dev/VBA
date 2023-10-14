VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColorClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ArrMain As Variant
Private ArrColor As Variant
Private DictKeys As New Dictionary
Private DictItems As New Dictionary
Private DictColor As New Dictionary
'==================================================================================================================
Private Sub Class_Initialize()
ArrColor = Array("3", "4", "5", "6", "7", "8", "12", "14", "16", "19", "23")
ArrMain = [DataRes]
End Sub
Public Sub StartColorClass()
Dim i&, CountColor%, key$: CountColor = -1
Dim Check As Boolean
Dim arrTmp As Variant

For i = LBound(ArrMain) To UBound(ArrMain)
    If ArrMain(i, 14) <> "р" Then
        key = ArrMain(i, 10) & "|" & ArrMain(i, 15)
        If Not DictKeys.Exists(key) Then
            DictKeys.Add key, i
        Else
            If Not DictColor.Exists(DictKeys(key)) Then
                CountColor = CountColor + 1
                If CountColor = 11 Then CountColor = 0
                DictColor.Add DictKeys(key), CountColor
            End If
            DictColor.Add i, CountColor
        End If
    End If
Next i

arrTmp = Array(DictColor.Keys, DictColor.Items)

[DataRes].Interior.ColorIndex = 0
For i = LBound(arrTmp(0)) To UBound(arrTmp(0))
    [DataRes].Columns(1).Resize(, 17).Rows(arrTmp(0)(i)).Interior.Color = 1550558
    [DataRes].Columns(18).Resize(, 2).Rows(arrTmp(0)(i)).Interior.ColorIndex = ArrColor(arrTmp(1)(i))
Next i

End Sub
Public Property Get arrResultsMain()
arrResultsMain = arrResMain
End Property



