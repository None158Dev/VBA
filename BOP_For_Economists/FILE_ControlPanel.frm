VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FILE_ControlPanel 
   Caption         =   "Панель управления"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   OleObjectBlob   =   "FILE_ControlPanel.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FILE_ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub PrintGroup_Click()
FILE_ControlPanel.Hide
Application.Run "DataGroup.Group"
FILE_ControlPanel.Show
End Sub
Private Sub PrintPaint_Click()
Application.Run "Func.FilterColorFormula"
End Sub
Private Sub PrintPercent_Click()
Application.Run "DataCount.CountPetcent"
End Sub
Private Sub PrintSvod_Click()
Application.Run "DataKeysUniq.SmUniqKeys"
End Sub
Private Sub UserForm_Initialize()
UniqFalse.Visible = False
ComplitTime.Visible = False
ComplitTimeSvod.Visible = False
ComplitTimePercent.Visible = False
ComplitTimeGroup.Visible = False
ComplitTimePaint.Visible = False
End Sub
Private Sub Cancel_Click()
Unload Me
End Sub
Private Sub EscClose_Click()
Unload Me
End Sub
Private Sub PrintSelect_Click()
Application.Run "DataUniq.Uniq"
End Sub
