VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HandlerArrayClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private NumberError%
Private ArrRange As Variant
Private ArrTotal As Variant
Private KeyGen$
Private arrResMain As Variant
Private arrResSec As Variant
Private DictMainKeys As New Dictionary
Private DictSum As New Dictionary
Private DictSumSmeta As New Dictionary
Private DictSecUniq As New Dictionary
'==================================================================================================================
Private Sub Class_Initialize()
KeyGen = Fix(Now * 10000000000#)
End Sub
Public Sub StartHandlerArray(ByVal MassiveRange As Variant, ByVal MassiveSvod As Variant)
Dim i&, G&
Dim SaveSection$: SaveSection = "—"
Dim SaveGroup$: SaveGroup = "—"
Dim Types$, NumbItem$, NumbSec$, NumbGr$, KeySmeta$, KeySection$, KeyGroup$, mesu$
Dim CheckPRDX As Boolean
Dim CheckGR As Boolean
Dim KeyJob As Long
Dim volume As Double
Dim arrTmp As Variant

ArrRange = MassiveRange
ArrTotal = MassiveSvod

For G = LBound(ArrRange) To UBound(ArrRange)
    For i = LBound(ArrRange(G)) To UBound(ArrRange(G))
        Call Check_Sec_Gr(ArrRange(G)(i, 1), ArrRange(G)(i, 3), SaveSection, SaveGroup, CheckPRDX, CheckGR, NumbSec)
        If ArrRange(G)(i, 2) <> "" And Not ArrRange(G)(i, 2) Like "*999-*" Then
        If ArrRange(G)(i, 1) <> "" Then NumbItem = ArrRange(G)(i, 1)
            KeySmeta = ArrTotal(G, 2)
            KeySection = KeysSection(ArrTotal(G, 2), NumbSec)
            KeyGroup = KeysGroup(KeySection, CheckPRDX, CheckGR, NumbGr)
            
            Call Mesure(ArrRange(G)(i, 4), ArrRange(G)(i, 5), mesu, volume)
            Call ClearMassive(ArrRange(G)(i, 2))
            
            If PRDX(ArrRange(G)(i, 2), CheckPRDX, Types, ArrRange(G)(i, 3)) Then
                NumbItem = ArrRange(G)(i, 1)
                KeyGen = KeyGen + 1
                DictMainKeys.Add DictMainKeys.Count + 1, Array(NumbItem, "PRDX", "р", "Вспомогательная расценка", "Р_" & KeyGen, KeySmeta, KeySection, KeyGroup, SaveSection, SaveGroup, "шт", 1, 1)
                CheckGR = False
            End If
            
            If Types = "р" Then
                KeyJob = DictMainKeys.Count
                If ArrRange(G)(i, 10) <> "" Then
                    DictSum(KeyJob) = DictSum(KeyJob) + CDbl(ArrRange(G)(i, 10))
                End If
            End If
            DictMainKeys.Add DictMainKeys.Count + 1, Array(NumbItem, ClearMassive(ArrRange(G)(i, 2)), Types, ClearMassive(ArrRange(G)(i, 3)), "Р_" & KeyGen, KeySmeta, KeySection, KeyGroup, SaveSection, SaveGroup, mesu, volume, ArrRange(G)(i, 10))
            If Not DictSecUniq.Exists(KeySmeta & "|" & NumbSec) Then DictSecUniq.Add KeySmeta & "|" & NumbSec, Array(KeySmeta, NumbSec, SaveSection)
        End If

        If ArrRange(G)(i, 3) = "ЗП" Or ArrRange(G)(i, 3) = "ЭМ" Or ArrRange(G)(i, 3) = "НР от ФОТ" Or ArrRange(G)(i, 3) = "СП от ФОТ" Or ArrRange(G)(i, 2) Like "*999-*" Then
            If ArrRange(G)(i, 10) <> "" Then
                DictSum(KeyJob) = DictSum(KeyJob) + CDbl(ArrRange(G)(i, 10))
            End If
        End If
    Next i
Next G

arrTmp = Array(DictMainKeys.Keys, DictMainKeys.Items, DictSecUniq.Items)

ReDim arrResMain(UBound(arrTmp(0)), 1 To 19)
ReDim arrResSec(UBound(arrTmp(2)), 1 To 3)

For i = LBound(arrTmp(0)) To UBound(arrTmp(0))
    arrResMain(i, 1) = arrTmp(1)(i)(0)
    arrResMain(i, 2) = "С_ТС"
    arrResMain(i, 3) = "—"
    arrResMain(i, 4) = arrTmp(1)(i)(5)
    arrResMain(i, 5) = arrTmp(1)(i)(6)
    arrResMain(i, 6) = arrTmp(1)(i)(7)
    arrResMain(i, 8) = arrTmp(1)(i)(8)
    arrResMain(i, 9) = arrTmp(1)(i)(9)
    arrResMain(i, 10) = arrTmp(1)(i)(4)
    arrResMain(i, 12) = "—"
    arrResMain(i, 14) = arrTmp(1)(i)(2)
    arrResMain(i, 15) = arrTmp(1)(i)(1)
    arrResMain(i, 16) = arrTmp(1)(i)(3)
    arrResMain(i, 17) = arrTmp(1)(i)(10)
    arrResMain(i, 18) = Round(arrTmp(1)(i)(11), 6)
    arrResMain(i, 19) = arrTmp(1)(i)(12)
    
    If DictSum.Exists(i) Then arrResMain(i, 19) = DictSum(i)
    If arrResMain(i, 19) = "" Then arrResMain(i, 19) = 0
    DictSumSmeta(arrResMain(i, 4)) = DictSumSmeta(arrResMain(i, 4)) + arrResMain(i, 19)
Next i

For i = LBound(ArrTotal) To UBound(ArrTotal)
    ArrTotal(i, 4) = DictSumSmeta(ArrTotal(i, 2))
    ArrTotal(i, 5) = ArrTotal(i, 3) - DictSumSmeta(ArrTotal(i, 2))
Next i

For i = LBound(arrTmp(2)) To UBound(arrTmp(2))
    arrResSec(i, 1) = arrTmp(2)(i)(0)
    arrResSec(i, 2) = arrTmp(2)(i)(1)
    arrResSec(i, 3) = arrTmp(2)(i)(2)
Next i

End Sub
Private Function Check_Sec_Gr(ByVal value As String, ByVal value2 As String, ByRef SaveSection As String, ByRef SaveGroup As String, ByRef CheckPRDX As Boolean, Optional ByRef CheckGR As Boolean, Optional NumbSec As String, Optional CheckSec As Boolean)
If value Like "*Раздел*" Then SaveSection = SectionSp(value, NumbSec): CheckSec = True: CheckGR = True
If Not IsNumeric(value) And Not value Like "*Раздел*" And value <> "" And value2 = "" Then SaveGroup = Application.WorksheetFunction.Trim(value): CheckPRDX = True
If CheckSec = True Then SaveGroup = "—": CheckPRDX = True
End Function
Private Function PRDX(ByVal value As String, ByRef CheckPRDX As Boolean, ByRef Types As String, ByVal value2 As String) As Boolean: Types = "м"
    If (Not value2 Like "*добавл*") And (value Like "*ФЕР*" Or value Like "*ТЕР*" Or value Like "*ТССЦпг*" Or value Like "*ФССЦпг*" Or value Like "*ФСЭМ*") Then
        PRDX = False
        CheckPRDX = False
        Types = "р"
        KeyGen = KeyGen + 1
    Else
        If CheckPRDX Then
            PRDX = True
            CheckPRDX = False
        Else
            PRDX = False
        End If
    End If
End Function
Private Function SectionSp(ByVal SaveSection As String, Optional ByRef NumbSec As String) As String
Dim arrSp As Variant
arrSp = Split(Replace(SaveSection, ".", Chr(164), , 1), Chr(164))
NumbSec = RegExp(arrSp(0), "(\D)")

If arrSp(1) = "" Or UBound(arrSp) = 0 Then
    SectionSp = "—"
Else
    SectionSp = Application.WorksheetFunction.Trim(arrSp(1))
End If
End Function
Private Function RegExp(ByVal value As String, ByVal Patt As String) As String
Dim myRegExp As New RegExp
With myRegExp
    .Global = True
    .Pattern = Patt
End With
    RegExp = myRegExp.Replace(value, "")
End Function
Private Function KeysSection(ByVal value As String, ByVal Numb As String) As String
If Len(Numb) = 1 Then KeysSection = value & "-00" & Numb
If Len(Numb) = 2 Then KeysSection = value & "-0" & Numb
If Len(Numb) = 3 Then KeysSection = value & "-" & Numb
End Function
Private Function KeysGroup(ByVal value As String, ByVal CheckPRDX As Boolean, ByRef CheckGR As Boolean, ByRef NumbGr As String) As String
If CheckGR Then NumbGr = 0: CheckGR = False
If CheckPRDX Then
    NumbGr = NumbGr + 1
    If Len(NumbGr) = 1 Then KeysGroup = value & "=00" & NumbGr
    If Len(NumbGr) = 2 Then KeysGroup = value & "=0" & NumbGr
    If Len(NumbGr) = 3 Then KeysGroup = value & "=" & NumbGr
Else
    If Len(NumbGr) = 1 Then KeysGroup = value & "=00" & NumbGr
    If Len(NumbGr) = 2 Then KeysGroup = value & "=0" & NumbGr
    If Len(NumbGr) = 3 Then KeysGroup = value & "=" & NumbGr
End If
End Function
Private Function Mesure(ByVal value As String, ByVal value2 As String, Optional ByRef mesu As String, Optional ByRef volume As Double)
Dim arrSp As Variant
If value = "" And value2 = "" Then Exit Function
If value2 = "" Then value2 = 1
arrSp = Split(value, " ")
If UBound(arrSp) > 1 Then arrSp = Split(Replace(value, " ", Chr(164), , 1), Chr(164)): mesu = arrSp(1): volume = arrSp(0) * value2
If UBound(arrSp) = 1 And IsNumeric(arrSp(0)) Then mesu = arrSp(1): volume = arrSp(0) * value2
If UBound(arrSp) = 0 Then mesu = value: volume = 1 * value2
If UBound(arrSp) = 1 And Not IsNumeric(arrSp(0)) Then mesu = value: volume = 1 * value2
End Function
Private Function ClearMassive(ByVal value As String) As String
If InStr(value, Chr(10)) - 1 <> -1 Then
    ClearMassive = Split(value, Chr(10))(0)
Else
    ClearMassive = value
End If

If ClearMassive Like "ФССЦ-*" Then
    ClearMassive = Replace(ClearMassive, "ФССЦ-", "")
ElseIf ClearMassive Like "ТССЦ-*" Then
    ClearMassive = Replace(ClearMassive, "ТССЦ-", "")
ElseIf ClearMassive Like "ФСЦ-*" Then
    ClearMassive = Replace(ClearMassive, "ФСЦ-", "")
ElseIf ClearMassive Like "ТСЦ-*" Then
    ClearMassive = Replace(ClearMassive, "ТСЦ-", "")
End If

End Function
Public Property Get arrResultsMain()
arrResultsMain = arrResMain
End Property
Public Property Get arrResultsTotal()
arrResultsTotal = ArrTotal
End Property
Public Property Get arrResultsSec()
arrResultsSec = arrResSec
End Property
