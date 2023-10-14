VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FILE_ControlPanel 
   Caption         =   "Панель управления"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   OleObjectBlob   =   "FILE_ControlPanel.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FILE_ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
Path.Text = [pos_all].Worksheet.Range("XFD1")
UniqFalse.Visible = False
SaveFalse.Visible = False
ComplitTime.Visible = False
End Sub
Private Sub Cancel_Click()
Unload Me
End Sub
Private Sub EscClose_Click()
Unload Me
End Sub
Private Sub PrintSelect_Click()
Application.Run "Data.DataMain"
End Sub
Private Sub SaveListAll_Click()
Application.Run "Data.SaveList"
End Sub
Private Sub SavePath_Click()
[pos_all].Worksheet.Range("XFD1") = Path.Text
ThisWorkbook.Save
End Sub
