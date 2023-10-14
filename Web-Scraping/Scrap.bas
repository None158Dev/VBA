Attribute VB_Name = "Scrap"
Option Explicit
Function GetHTTPS(ByVal URL As String) As String

With CreateObject("msxml2.xmlhttp")
    .Open "GET", URL, False
    .send
Do: DoEvents: Loop Until .readyState = 4
    GetHTTPS = .responseText
End With

End Function
Function RegEx(MyRange As Range, MyReplace As String) As String
Dim reg_exp As New RegExp

With reg_exp
    .Global = True
    .IgnoreCase = True
    .Pattern = "\D"
End With
    RegEx = reg_exp.Replace(MyRange, MyReplace)
    
End Function
Function TransposeArr(ByVal MyArr As Variant) As Variant
Dim i&, G&
Dim ArrTemp As Variant
ReDim ArrTemp(LBound(MyArr, 1) To UBound(MyArr, 1), 0 To UBound(MyArr(0)))
 
    For i = LBound(MyArr, 1) To UBound(MyArr, 1)
        For G = 0 To UBound(MyArr(0))
            ArrTemp(i, G) = MyArr(i)(G)
        Next G
    Next i
     
TransposeArr = ArrTemp
End Function
Private Sub Catalog()
Dim i&, G&, J&
Dim URL As String
Dim GetUrl As String
Dim Infinity As Double
'==========================================
Dim docHTML As Object
Dim getCatalogUrl As Object
'==========================================
Dim Razdel As String
Dim Group As String
Dim StrURL As String
'==========================================
Dim ArrUrl() As Variant
'==========================================
Dim dic_CatalogURL As New Dictionary

Infinity = Timer
ObjectsUpdate
Application.ScreenUpdating = False
URL = "https://www.maxidom.ru/"
Set docHTML = CreateObject("HTMLFILE")
GetUrl = GetHTTPS(URL)
docHTML.body.innerHTML = GetUrl



Set getCatalogUrl = docHTML.getElementsByClassName("content catalog-wrap")(0).Children(0)

For i = 0 To getCatalogUrl.Children(1).Children.Length - 1

    Razdel = getCatalogUrl.Children(0).Children(i).Children(0).Children(1).innerText
    
    For G = 0 To getCatalogUrl.Children(1).Children(i).Children.Length - 1
        For J = 0 To getCatalogUrl.Children(1).Children(i).Children(G).Children.Length - 1

            If getCatalogUrl.Children(1).Children(i).Children(G).Children(J).className = "menu-catalog__item-lvl2 menu-catalog__item-lvl2-h" Then
                Group = WorksheetFunction.Trim(getCatalogUrl.Children(1).Children(i).Children(G).Children(J).Children(0).innerText)
                StrURL = Replace(getCatalogUrl.Children(1).Children(i).Children(G).Children(J).Children(0).href, "about:/", URL) & "?amount=5000"
                dic_CatalogURL.Add dic_CatalogURL.Count, Array(Razdel, Group, StrURL)
            End If
        Next J
    Next G
Next i

ArrUrl = TransposeArr(dic_CatalogURL.Items)

st_URL.Range(2, 1).Resize(UBound(ArrUrl) + 1, UBound(Application.Transpose(ArrUrl))) = ArrUrl
Application.ScreenUpdating = True
MsgBox "Ссылки забраны!", vbInformation, Format(Timer - Infinity, "0.00 сек")
End Sub
Private Sub AllPos()
Dim i&, G&, J&
Dim URL As String
Dim GetUrl As String
Dim siteMain As String
Dim Infinity As Double
'=========================================
Dim Name As String
Dim UrlPos As String
Dim Art As String
Dim KeyPos As String
Dim Brand As String
Dim Country As String
Dim Weight As Double
Dim Price As String
Dim Measure As String
Dim Razdel As String
Dim Group As String
Dim Subgroup As String
'=========================================
Dim docHTML As Object
Dim getBlock As Object
Dim getBlockName As Object
Dim getBlockPrice As Object
Dim getBlockCatalog As Object
'=========================================
Dim SaveURL_SubGroup As String
'=========================================
Dim ArrData() As Variant
Dim ArrUrl() As Variant
'=========================================
Dim dic_CatalogPos As New Dictionary

Infinity = Timer

ObjectsUpdate
ArrUrl = [_UrlAdd].Value2
URL = "https://www.maxidom.ru/"
Set docHTML = CreateObject("HTMLFILE")
GetUrl = GetHTTPS(URL)
docHTML.body.innerHTML = GetUrl
'=========================================
'Номеклатура, Url, Арт, Код товара, Марка, Страна, Вес, Цена, Ед.Изм, Раздел, Группа, Подгруппа
On Error Resume Next:
For i = 0 To UBound(ArrUrl)
'UBound(ArrUrl)
Application.Wait (Now + TimeValue("0:00:01") / 2)
GetUrl = GetHTTPS(ArrUrl(i + 1, 1))
docHTML.body.innerHTML = GetUrl
Set getBlock = docHTML.getElementsByTagName("article")
Cells(1, 13) = i
    For G = 0 To getBlock.Length - 1
        
        Set getBlockName = getBlock(G).getElementsByClassName("b-catalog-list-product__section-left")(0). _
        getElementsByClassName("b-catalog-list-product__section2 caption-list")(0)
        
        If getBlockName.Children.Length <> 0 Then
        
            Name = WorksheetFunction.Trim(getBlockName.Children(0).innerText)
            UrlPos = Replace(getBlockName.Children(0).href, "about:/", URL)
            Art = getBlockName.Children(1).Children(0).Children(1).innerText
            KeyPos = getBlockName.Children(1).Children(1).Children(1).innerText
            Brand = getBlockName.Children(1).Children(2).Children(1).innerText
            Country = getBlockName.Children(2).Children(0).innerText
            Weight = Replace(Split(getBlockName.Children(2).Children(1).innerText, " ")(0), ".", ",")
            
            Set getBlockPrice = getBlock(G).getElementsByClassName("price-list")(0)
            
            Price = Replace(getBlockPrice.Children(0).innerText, " ", "")
            Measure = WorksheetFunction.Trim(Replace(getBlockPrice.Children(4).innerText, ".", ""))
            
            Set getBlockCatalog = docHTML.getElementsByClassName("breadcrumbs__ol")(0)
            
            Razdel = getBlockCatalog.Children(2).Children(0).innerText
            Group = getBlockCatalog.Children(3).Children(0).innerText
            
            dic_CatalogPos.Add dic_CatalogPos.Count, Array(Razdel, Group, Name, Price, Measure, Weight, Brand, Country, Art, KeyPos, UrlPos)
        End If
    Next G
Next i
On Error GoTo 0
ArrData = TransposeArr(dic_CatalogPos.Items)

st_Pos.Range(2, 1).Resize(UBound(ArrData) + 1, UBound(Application.Transpose(ArrData))) = ArrData
MsgBox "Позиции забраны!", vbInformation, Format(Timer - Infinity, "0.00 сек")
End Sub
