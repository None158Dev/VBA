VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Name class module: MyGroup
Private MainArr As Variant
Private dicSm As New Dictionary
Private dicSection As New Dictionary
Private dicGroup As New Dictionary
Private dicJob As New Dictionary
Private Sub Class_Initialize()
MainArr = [BopSebes].Value2
End Sub
Private Function Smeta(ByVal index As Long, _
                       ByRef UBoundP_Sm As Long, _
                       Optional ByRef LBoundP_Sm As Long, _
                       Optional ByRef CountSm As Long)

CountSm = CountSm + 1

If CountSm = 1 Then
    LBoundP_Sm = index + 1
ElseIf CountSm = 2 Then
    dicSm.Add LBoundP_Sm, UBoundP_Sm
    LBoundP_Sm = index + 1
    CountSm = 1
End If

End Function
Private Function Section(ByVal index As Long, _
                       ByRef UBoundP_Sec As Long, _
                       Optional ByRef LBoundP_Sec As Long, _
                       Optional ByRef CountSec As Long)

CountSec = CountSec + 1

If CountSec = 1 Then
    LBoundP_Sec = index + 1
ElseIf CountSec = 2 Then
    dicSection.Add LBoundP_Sec, UBoundP_Sec
    LBoundP_Sec = index + 1
    CountSec = 1
End If

End Function
Private Function Group(ByVal index As Long, _
                       ByRef UBoundP_GR As Long, _
                       Optional ByRef LBoundP_GR As Long, _
                       Optional ByRef CountGR As Long)
CountGR = CountGR + 1

If CountGR = 1 Then
    LBoundP_GR = index + 1
ElseIf CountGR = 2 Then
    dicGroup.Add LBoundP_GR, UBoundP_GR
    LBoundP_GR = index + 1
    CountGR = 1
End If

End Function
Public Sub Start()
Dim i&, CountPos&, LBoundP&, UBoundP&, CountP&, CountNotP&, TempPosP&, TempPosI&
Dim LBoundP_GR&, UBoundP_GR&, CountGR&
Dim LBoundP_Sec&, UBoundP_Sec&, CountSec&
Dim LBoundP_Sm&, UBoundP_Sm&, CountSm&
Dim ArrJob As Variant
Dim ArrGroup As Variant
Dim ArrSection As Variant
Dim ArrSmeta As Variant

On Error Resume Next
For i = LBound(MainArr) To UBound(MainArr): CountPos = CountPos + 1
    If MainArr(i, 7) = "Смета" Then
        Call ResetCounts(CountPos, CountNotP, CountP)
        If CountNotP = 1 Then
            UBoundP = i - 1
            dicJob.Add LBoundP, UBoundP
        End If
    ElseIf MainArr(i, 7) = "Раздел" Then
        Call ResetCounts(CountPos, CountNotP, CountP)

        If CountNotP = 1 Then
            UBoundP = i - 1
            dicJob.Add LBoundP, UBoundP
        End If
    ElseIf MainArr(i, 7) = "Группа" Then
        Call ResetCounts(CountPos, CountNotP, CountP)
        If CountNotP = 1 Then
            UBoundP = i - 1
            dicJob.Add LBoundP, UBoundP
        End If
    ElseIf MainArr(i, 7) = "р" Then
        CountPos = 0
        CountNotP = 0
        CountP = CountP + 1
        TempPosP = i
        If TempPosP - TempPosI = 1 Then
            CountP = 1
        End If
    End If

TempPosI = TempPosP

    If CountPos = 1 Then
        LBoundP = i
    ElseIf CountPos = 0 Then
        If CountP = 2 Then
            UBoundP = i - 1
            CountP = 1
            dicJob.Add LBoundP, UBoundP
        End If
    End If
'=======================================================================
    If MainArr(i, 7) = "Группа" Then
        Call Group(i, UBoundP, LBoundP_GR, CountGR)
    End If
'=======================================================================
    If MainArr(i, 7) = "Раздел" Then
        Call Section(i, UBoundP, LBoundP_Sec, CountSec)
    End If
'=======================================================================
    If MainArr(i, 7) = "Смета" Then
        Call Smeta(i, UBoundP, LBoundP_Sm, CountSm)
    End If
'=======================================================================
    If i = UBound(MainArr) Then
        dicJob.Add LBoundP, i
        dicGroup.Add LBoundP_GR, i
        dicSection.Add LBoundP_Sec, i
        dicSm.Add LBoundP_Sm, i
    End If

Next i
On Error GoTo 0

dicJob.Remove (0)

ArrJob = Array(dicJob.Keys, dicJob.Items)
ArrGroup = Array(dicGroup.Keys, dicGroup.Items)
ArrSection = Array(dicSection.Keys, dicSection.Items)
ArrSmeta = Array(dicSm.Keys, dicSm.Items)

For i = 0 To UBound(ArrJob(0))
    [BopSebes].Rows(ArrJob(0)(i) & ":" & ArrJob(1)(i)).Group
Next i

For i = 0 To UBound(ArrGroup(0))
    [BopSebes].Rows(ArrGroup(0)(i) & ":" & ArrGroup(1)(i)).Group
Next i

For i = 0 To UBound(ArrSection(0))
    [BopSebes].Rows(ArrSection(0)(i) & ":" & ArrSection(1)(i)).Group
Next i

For i = 0 To UBound(ArrSmeta(0))
    [BopSebes].Rows(ArrSmeta(0)(i) & ":" & ArrSmeta(1)(i)).Group
Next i

End Sub
Private Function ResetCounts(ByRef CountPos As Long, _
                             ByRef CountNotP As Long, _
                             ByRef CountP As Long)
CountPos = 0
CountNotP = CountNotP + 1
CountP = 0
End Function
Public Property Get UniqItems() As Variant
UniqItems = dicJob.Items
End Property
Public Property Get UniqKeys() As Variant
UniqKeys = dicJob.Keys
End Property
