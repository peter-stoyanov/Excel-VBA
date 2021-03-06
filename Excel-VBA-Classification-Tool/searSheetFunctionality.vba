Private Sub CLASSIFBTN_Click()


With Application
.ScreenUpdating = False
.Calculation = xlCalculationManual
.DisplayStatusBar = False
End With

Dim search As Worksheet

Dim lrsearch As Long
Dim i As Long
Dim printrow As Long
Dim filterTerm As String


Set search = ThisWorkbook.Worksheets("search")

filterTerm = search.Cells(2, 6).Value

lr_search1 = search.Cells(search.Rows.Count, "CX").End(xlUp).row
lr_search2 = search.Cells(search.Rows.Count, "DC").End(xlUp).row


Dim search1() As String
ReDim search1(1 To lr_search1)
For i = 10 To lr_search1
        search1(i) = CStr(search.Cells(i, 100)) + "  |  " + CStr(search.Cells(i, 101)) + "  |  " + CStr(search.Cells(i, 102)) + "  |  " + CStr(search.Cells(i, 103))
Next

Dim search2() As String
ReDim search2(1 To lr_search2)
For i = 10 To lr_search2
        search2(i) = CStr(search.Cells(i, 105)) + "  |  " + CStr(search.Cells(i, 106)) + "  |  " + CStr(search.Cells(i, 107)) + "  |  " + CStr(search.Cells(i, 108)) + "  |  " + CStr(search.Cells(i, 109)) + "  |  " + CStr(search.Cells(i, 110))
Next

search.Range("A7:J10000").Clear
search.Range("A7:J10000").NumberFormat = "@"

'search
printrow = 8
Dim freecol As Long


Dim splitarr As Variant

'--------------------------classification search
search.Cells(printrow, 3) = "Results from Classifications"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(search1)
    freecol = 1
    If InStr(1, search1(i), filterTerm, vbTextCompare) > 0 Then
        splitarr = Split(search1(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol) = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

'--------------------------types search
printrow = 8
search.Cells(printrow, 6) = "Results from Existing Connected Product Types"
search.Cells(printrow, 6).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(search2)
    freecol = 5
    If InStr(1, search2(i), filterTerm, vbTextCompare) > 0 Then
        splitarr = Split(search2(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol) = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'----------------------




 With search.Range("A7:J10000")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With

With Application
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic
.DisplayStatusBar = True
.Calculate
End With

End Sub

Private Sub FilterBtn_Click()

With Application
.ScreenUpdating = False
.Calculation = xlCalculationManual
.DisplayStatusBar = False
End With

Dim search As Worksheet

Dim lrsearch As Long
Dim i As Long
Dim printrow As Long
Dim filterTerm As String


Set search = ThisWorkbook.Worksheets("search")

filterTerm = search.Cells(4, 3).Value

lr_search1 = search.Cells(search.Rows.Count, "CX").End(xlUp).row
lr_search2 = search.Cells(search.Rows.Count, "DC").End(xlUp).row


Dim search1() As String
ReDim search1(1 To lr_search1)
For i = 10 To lr_search1
        search1(i) = CStr(search.Cells(i, 100)) + "  |  " + CStr(search.Cells(i, 101)) + "  |  " + CStr(search.Cells(i, 102)) + "  |  " + CStr(search.Cells(i, 103))
Next

Dim search2() As String
ReDim search2(1 To lr_search2)
For i = 10 To lr_search2
        search2(i) = CStr(search.Cells(i, 105)) + "  |  " + CStr(search.Cells(i, 106)) + "  |  " + CStr(search.Cells(i, 107)) + "  |  " + CStr(search.Cells(i, 108)) + "  |  " + CStr(search.Cells(i, 109)) + "  |  " + CStr(search.Cells(i, 110))
Next

search.Range("A7:J10000").Clear
search.Range("A7:J10000").NumberFormat = "@"

'search
printrow = 8
Dim freecol As Long


Dim splitarr As Variant

'--------------------------classification search
search.Cells(printrow, 3) = "Results from Classifications"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(search1)
    freecol = 1
    If InStr(1, search1(i), filterTerm, vbTextCompare) > 0 Then
        splitarr = Split(search1(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol) = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

'--------------------------types search
printrow = 8
search.Cells(printrow, 6) = "Results from Existing Connected Product Types"
search.Cells(printrow, 6).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(search2)
    freecol = 5
    If InStr(1, search2(i), filterTerm, vbTextCompare) > 0 Then
        splitarr = Split(search2(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol) = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'----------------------




 With search.Range("A7:J10000")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With

With Application
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic
.DisplayStatusBar = True
.Calculate
End With




End Sub

Private Sub SearchBtn_Click()
'this is the initial search

With Application
.ScreenUpdating = False
.Calculation = xlCalculationManual
.DisplayStatusBar = False
End With



Dim types As Worksheet
Dim clasif As Worksheet

Dim lrsearch As Long
Dim i As Long
Dim printrow As Long
Dim searchTerm As String
'Dim filterTerm As String



Set search = ThisWorkbook.Worksheets("search")
Set types = ThisWorkbook.Worksheets("types")
Set clasif = ThisWorkbook.Worksheets("classifications")

'here clear the hidden temp place for search results
search.Range("CV7:DF10000").ClearContents
search.Range("CV7:DF10000").NumberFormat = "@"

searchTerm = search.Cells(2, 3).Value
'filterTerm = search.Cells(5, 3).Value

search.Range("A7:J10000").Clear
search.Range("A7:J10000").NumberFormat = "@"

lr_types = types.Cells(types.Rows.Count, "C").End(xlUp).row
lr_clasif = clasif.Cells(clasif.Rows.Count, "C").End(xlUp).row


Dim arrClassif() As String
ReDim arrClassif(1 To lr_clasif)
For i = 2 To lr_clasif
        'Var = clasif.Cells(i, 1).Value
        arrClassif(i) = CStr(clasif.Cells(i, 1)) + "  |  " + CStr(clasif.Cells(i, 2)) + "  |  " + CStr(clasif.Cells(i, 3)) + "  |  " + CStr(clasif.Cells(i, 19))
Next

Dim arrTypes() As String
ReDim arrTypes(1 To lr_types)
For i = 2 To lr_types
        arrTypes(i) = CStr(types.Cells(i, 1)) + "  |  " + CStr(types.Cells(i, 2)) + "  |  " + CStr(types.Cells(i, 3)) + "  |  " + CStr(types.Cells(i, 4)) + "  |  " + CStr(types.Cells(i, 5)) + "  |  " + CStr(types.Cells(i, 6))
Next


'search
printrow = 8
Dim freecol As Long


Dim splitarr As Variant

'--------------------------classification search
search.Cells(printrow, 3) = "Results from Classifications"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrClassif)
    'If printrow = 13 Then Stop
    freecol = 1
    freecol2 = 100
    If InStr(1, arrClassif(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrClassif(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    If Len(splitarr(j)) > 2 Then
                        search.Cells(printrow, freecol) = splitarr(j)
                        search.Cells(printrow, freecol2) = splitarr(j)
                        freecol = freecol + 1
                        freecol2 = freecol2 + 1
                    End If
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

'--------------------------types search
printrow = 8
search.Cells(printrow, 6) = "Results from Existing Connected Product Types"
search.Cells(printrow, 6).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrTypes)
    freecol = 5
    freecol2 = 105
    If InStr(1, arrTypes(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrTypes(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol) = splitarr(j)
                    search.Cells(printrow, freecol2) = splitarr(j)
                    freecol = freecol + 1
                    freecol2 = freecol2 + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'----------------------




 With search.Range("A7:J10000")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With

With Application
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic
.DisplayStatusBar = True
.Calculate
End With


End Sub


Sub Gotocode()

Dim rowToGo As Long
Dim row As Long
Dim searchSht As Worksheet

Set searchSht = ThisWorkbook.Worksheets("search")
row = ActiveCell.row
searchSht.Cells(1, 100) = row
rowToGo = searchSht.Cells(row, 1)
If rowToGo > 0 Then
    
    ThisWorkbook.Worksheets("classifications").Activate
    ThisWorkbook.Worksheets("classifications").Cells(rowToGo, 3).Select
Else: MsgBox "Òðÿáâà äà ñè ñòúïèë íà ðåä ñ êîíêðåòåí ðåçóëòàò îò òúðñåíå â êîëîíà À"
End If


End Sub


Sub GoBack()

Dim rowBack As Long
Dim row As Long
Dim searchSht As Worksheet

Set searchSht = ThisWorkbook.Worksheets("search")

rowBack = searchSht.Cells(1, 100)

searchSht.Activate
searchSht.Cells(rowBack, 3).Select

searchSht.Cells(1, 100).Clear


End Sub

Sub Shop()

Dim rowActive As Long
Dim colActive As Long
'Dim row As Long
'Dim searchSht As Worksheet

Set searchSht = ThisWorkbook.Worksheets("search")
rowActive = ActiveCell.row
colActive = ActiveCell.Column

If colActive > 4 Then
    MsgBox "Çà äà ñè ñëîæèòå íåùî â òîðáè÷êàòà òðÿáâà äà ñòå ñòúïèëè íà êëåòêà ñ ðåçóëòàò îò òúðñåíåòî - êîëîíè îò À äî D"
Else: lr_shop = searchSht.Cells(searchSht.Rows.Count, "L").End(xlUp).row + 1
        searchSht.Cells(lr_shop, 12).Value = searchSht.Cells(rowActive, 2)
         searchSht.Cells(lr_shop, 13).Value = searchSht.Cells(rowActive, 3)
          searchSht.Cells(lr_shop, 14).Value = searchSht.Cells(rowActive, 4)
End If

End Sub

