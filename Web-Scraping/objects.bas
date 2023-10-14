Attribute VB_Name = "objects"
Option Explicit
Public sh_URL As Worksheet
Public sh_Pos As Worksheet

Public st_URL As ListObject
Public st_Pos As ListObject
Sub ObjectsUpdate()
Set sh_URL = [_Catalog].Worksheet
Set sh_Pos = [_Pos].Worksheet
Set st_URL = [_Catalog].ListObject
Set st_Pos = [_Pos].ListObject
End Sub
