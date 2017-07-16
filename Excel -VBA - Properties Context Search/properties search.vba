Private Sub CommandButton1_Click()


With Application
.ScreenUpdating = False
.Calculation = xlCalculationManual
.DisplayStatusBar = False
End With

Dim TS As Worksheet
Dim EC As Worksheet
Dim P As Worksheet
Dim Ref As Worksheet
Dim M As Worksheet
Dim U As Worksheet
Dim V As Worksheet
Dim IU As Worksheet
Dim PS As Worksheet
Dim search As Worksheet

Dim lrTS As Long
Dim lrEC As Long
Dim lrP As Long
Dim lrRef As Long
Dim lrM As Long
Dim lrU As Long
Dim lrV As Long
Dim lrIU As Long
Dim lrPS As Long
Dim lrsearch As Long
Dim i As Long
Dim printrow As Long
Dim searchTerm As String
Dim filter As String


Set TS = ActiveWorkbook.Worksheets("Technical Specifications")
Set EC = ActiveWorkbook.Worksheets("Essential Characteristics")
Set P = ActiveWorkbook.Worksheets("Properties")
Set Ref = ActiveWorkbook.Worksheets("Reference Standards")
Set M = ActiveWorkbook.Worksheets("Measures")
Set U = ActiveWorkbook.Worksheets("Units")
Set V = ActiveWorkbook.Worksheets("Values")
Set IU = ActiveWorkbook.Worksheets("Intended Uses")
Set PS = ActiveWorkbook.Worksheets("Property Sets")
Set search = ActiveWorkbook.Worksheets("search")

searchTerm = search.Cells(2, 3).Value
filter = search.Cells(4, 3).Value

search.Range("A6:Z10000").Clear


lrTS = TS.Cells(TS.Rows.Count, "B").End(xlUp).Row
lrEC = EC.Cells(TS.Rows.Count, "B").End(xlUp).Row
lrP = P.Cells(TS.Rows.Count, "B").End(xlUp).Row
lrRef = Ref.Cells(TS.Rows.Count, "B").End(xlUp).Row
lrM = M.Cells(TS.Rows.Count, "B").End(xlUp).Row
lrU = U.Cells(TS.Rows.Count, "B").End(xlUp).Row
lrV = V.Cells(TS.Rows.Count, "B").End(xlUp).Row
lrIU = IU.Cells(TS.Rows.Count, "B").End(xlUp).Row
lrPS = PS.Cells(TS.Rows.Count, "B").End(xlUp).Row

Dim arrTS() As String
ReDim arrTS(1 To lrTS)
For i = 1 To lrTS
        arrTS(i) = TS.Cells(i, 2) + "  |  " + TS.Cells(i, 3) + "  |  " + TS.Cells(i, 4) + "  |  " + TS.Cells(i, 5)
Next

Dim arrEC() As String
ReDim arrEC(1 To lrEC)
For i = 1 To lrEC
        arrEC(i) = EC.Cells(i, 2) + "  |  " + EC.Cells(i, 3) + "  |  " + EC.Cells(i, 4) + "  |  " + EC.Cells(i, 5)
Next

Dim arrP() As String
ReDim arrP(1 To lrP)
For i = 1 To lrP
        arrP(i) = P.Cells(i, 2) + "  |  " + P.Cells(i, 3) + "  |  " + P.Cells(i, 4) + "  |  " + P.Cells(i, 5)
Next

Dim arrRef() As String
ReDim arrRef(1 To lrRef)
For i = 1 To lrRef
        arrRef(i) = Ref.Cells(i, 2) + "  |  " + Ref.Cells(i, 3) + "  |  " + Ref.Cells(i, 4) + "  |  " + Ref.Cells(i, 5)
Next

Dim arrM() As String
ReDim arrM(1 To lrM)
For i = 1 To lrM
        arrM(i) = M.Cells(i, 2) + "  |  " + M.Cells(i, 3) + "  |  " + M.Cells(i, 4) + "  |  " + M.Cells(i, 5)
Next

Dim arrU() As String
ReDim arrU(1 To lrU)
For i = 1 To lrU
        arrU(i) = U.Cells(i, 2) + "  |  " + U.Cells(i, 3) + "  |  " + U.Cells(i, 4) + "  |  " + U.Cells(i, 5)
Next

Dim arrV() As String
ReDim arrV(1 To lrV)
For i = 1 To lrV
        arrV(i) = V.Cells(i, 2) + "  |  " + V.Cells(i, 3) + "  |  " + V.Cells(i, 4) + "  |  " + V.Cells(i, 5)
Next

Dim arrIU() As String
ReDim arrIU(1 To lrIU)
For i = 1 To lrIU
        arrIU(i) = IU.Cells(i, 2) + "  |  " + IU.Cells(i, 3) + "  |  " + IU.Cells(i, 4) + "  |  " + IU.Cells(i, 5)
Next

Dim arrPS() As String
ReDim arrPS(1 To lrPS)
For i = 1 To lrPS
        arrPS(i) = PS.Cells(i, 2) + "  |  " + PS.Cells(i, 3) + "  |  " + PS.Cells(i, 4) + "  |  " + PS.Cells(i, 5)
Next

'search
printrow = 6
Dim freecol As Long


Dim splitarr As Variant

'--------------------------
search.Cells(printrow, 3) = "Technical Specifications"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrTS)
    freecol = 3
    If InStr(1, arrTS(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrTS(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol).Value = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

search.Cells(printrow, 3) = "Essential characteristics"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrEC)
    freecol = 3
    If InStr(1, arrEC(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrEC(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol).Value = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

search.Cells(printrow, 3) = "Properties"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrP)
    freecol = 3
    If InStr(1, arrP(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrP(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol).Value = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

search.Cells(printrow, 3) = "Reference standards"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrRef)
    freecol = 3
    If InStr(1, arrRef(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrRef(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol).Value = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

search.Cells(printrow, 3) = "Measures"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrM)
    freecol = 3
    If InStr(1, arrM(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrM(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol).Value = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

search.Cells(printrow, 3) = "Units"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrU)
    freecol = 3
    If InStr(1, arrU(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrU(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol).Value = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------
search.Cells(printrow, 3) = "Values"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrV)
    freecol = 3
    If InStr(1, arrV(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrV(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol).Value = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

search.Cells(printrow, 3) = "Intended uses"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrIU)
    freecol = 3
    If InStr(1, arrIU(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrIU(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol).Value = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

search.Cells(printrow, 3) = "Property Sets"
search.Cells(printrow, 3).Style = "Accent1"
printrow = printrow + 2
For i = 1 To UBound(arrPS)
    freecol = 3
    If InStr(1, arrPS(i), searchTerm, vbTextCompare) > 0 Then
        splitarr = Split(arrPS(i), "  |  ")
                For j = LBound(splitarr) To UBound(splitarr)
                    search.Cells(printrow, freecol).Value = splitarr(j)
                    freecol = freecol + 1
                Next j
        printrow = printrow + 1
    End If
Next i
printrow = printrow + 1
'--------------------------

 With search.Range("A6:Z10000")
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
