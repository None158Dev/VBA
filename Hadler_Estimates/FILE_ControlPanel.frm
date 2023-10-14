VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FILE_ControlPanel 
   Caption         =   "Панель управления"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   OleObjectBlob   =   "FILE_ControlPanel.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FILE_ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub PrintColor_Click()
Application.Run "Color.Main"
End Sub
Private Sub PrintDelSh_Click()
Application.Run "Func.Sheets_Del"
End Sub
Private Sub PrintHandler_Click()
Application.Run "Handler.Main"
End Sub
Private Sub PrintKeys_Click()
Application.Run "KeysPos.Main"
End Sub
Private Sub PrintUnloadSm_Click()
Application.Run "Func.importfile"
End Sub
Private Sub UserForm_Initialize()
ComplitTimeKeys.Visible = False
ComplitTimeColor.Visible = False
ComplitTimeHandler.Visible = False
End Sub
Private Sub Cancel_Click()
Unload Me
End Sub
Private Sub EscClose_Click()
Unload Me
End Sub
