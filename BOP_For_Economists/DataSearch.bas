Attribute VB_Name = "DataSearch"
Option Explicit
Private Sub Search()
Dim Infinity As Double
Dim shSeb As ListObject
Dim clsDict As MySearch

Call DisabledApps(False, False)
Infinity = Timer
[pos_UniqP].Columns(14).Calculate
[pos_UniqM].Columns(14).Calculate
Set clsDict = New MySearch

Set shSeb = [BopSebes].Worksheet.ListObjects(1): shSeb.AutoFilter.ShowAllData

clsDict.Start

[_Market].Resize(UBound(clsDict.Market), 1) = clsDict.Market
[_Status].Resize(UBound(clsDict.Status), 1) = clsDict.Status
Call DisabledApps(True, True)

MsgBox "Готово!", vbInformation, Format(Timer - Infinity, "0.00 сек")
End Sub


