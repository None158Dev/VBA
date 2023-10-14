VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Search_Form 
   Caption         =   "Поиск и выбор значений из списка:"
   ClientHeight    =   9645.001
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9555.001
   OleObjectBlob   =   "Search_Form.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Search_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub EscClose_Click()
    Unload Me
End Sub
Private Sub PDFSelect_Click()
Dim i&
Dim dic_PDF As New Dictionary
Dim reg_Clear As New RegExp
Dim arr As Variant
'===========================================================================
Application.ScreenUpdating = False
With reg_Clear
    .Global = True
    .IgnoreCase = True
    .Pattern = "[\]\[\n\r\{\}\^\.\$\(\)\+\?\*\|\-\\\!@#%&_№;%:-='/<>`~,”""]"
End With
'===========================================================================
For i = 0 To dic_UniqAct.Count - 1
    If ListArr.Selected(i) Then
        dic_PDF.Add ListArr.List(i, 0), i + 1
    End If
Next i
'===========================================================================
arr = dic_PDF.Items
'===========================================================================
For i = 1 To dic_PDF.Count
    [_VPR].Value2 = arr(i - 1)
    sh_Act.ExportAsFixedFormat Type:=xlTypePDF, Filename:=[_Path].Value2 & "\АКТ № " & reg_Clear.Replace([_NumberActB].Value2, "") & " от " & Format([_DataActB], "mm.dd.yy"), _
    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
Next i
'===========================================================================
Application.ScreenUpdating = True
End Sub
Private Sub PrintSelect_Click()
Dim i&
Dim dic_Print As New Dictionary
Dim arr As Variant
'===========================================================================
Application.ScreenUpdating = False
For i = 0 To dic_UniqAct.Count - 1
    If ListArr.Selected(i) Then
        dic_Print.Add ListArr.List(i, 0), i + 1
    End If
Next i
'===========================================================================
arr = dic_Print.Items
'===========================================================================
For i = 1 To dic_Print.Count
    [_VPR].Value2 = arr(i - 1)
    sh_Act.PrintOut
Next i
'===========================================================================
Application.ScreenUpdating = True
End Sub
Private Sub SavePath_Click()
Application.Run "Main.Path"
End Sub
Private Sub Search_DATA_Change()
ArrFilter = Filter(dic_UniqAct.Keys, Search_DATA.Value, True, vbTextCompare)
ListArr.List = ArrFilter
End Sub
Private Sub UserForm_Initialize()
'===========================================================================
Application.Run "Main.PDFUnload"
'===========================================================================
ListArr.List = dic_UniqAct.Keys
ListArr.MultiSelect = fmMultiSelectExtended
End Sub
