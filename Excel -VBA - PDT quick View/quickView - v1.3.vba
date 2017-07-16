Sub readandwrite()
'
' read data from batch and collect in quick compact view
'

Dim qVsht As Worksheet, batchsht As Worksheet, psht As Worksheet
Dim lastRow As Long
Dim freerow As Long
Dim freecol As Long
Dim j As Long
Dim i As Long
Dim propPsets As Collection
Dim it As Variant
Dim str As String
Dim lastrowPset As Long
Dim k As Long
Dim prop As String
Dim nest As String
Dim lastrB As Long
Dim lastrD As Long
Dim lastr As Long
Dim m As Long


Set qVsht = ActiveWorkbook.Worksheets("PDTquickView")
qVsht.Range("A2:P10000").Clear
'qVsht.Range("A2:M10000").Select
    With qVsht.Range("A2:P10000")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        '.Rows.AutoFit
        '.Columns.AutoFit
    End With
    
Set batchsht = ActiveWorkbook.Worksheets("paste")

lastRow_batch = batchsht.Cells(batchsht.Rows.Count, "C").End(xlUp).Row + 50
freerow = 1
freecol = 8

Set psht = ActiveWorkbook.Worksheets("psets")
'lastrowPset = psht.Cells(psht.Rows.Count, "A").End(xlUp).Row
lastrowPset = 4000
'Dim arr() As String
'ReDim arr(1 To lastrowPset)
'For i = 1 To lastrowPset
'        arr(i) = psht.Cells(i, 2)
'Next


For i = 112 To lastRow_batch
    If batchsht.Cells(i, 1) = "Property" Then
        If Len(batchsht.Cells(i, 3)) > 0 Then
            freerow = freerow + 1
            Set propPsets = New Collection
            qVsht.Cells(freerow, 3) = batchsht.Cells(i, 3)
            prop = qVsht.Cells(freerow, 3)
            qVsht.Cells(freerow, 3).Style = "Accent5"
            str = ""
            nest = ""
            For j = 5 To 205
                If batchsht.Cells(i, j) = "X" Then
                    propPsets.Add (j)
                    qVsht.Cells(freerow, freecol) = batchsht.Cells(3, j)
                    nest = qVsht.Cells(freerow, freecol).Value
                    If InStr(1, nest, "ifc", vbTextCompare) > 0 Then
                        If Not InStr(1, nest, "ifc prop", vbTextCompare) > 0 Then
                            qVsht.Cells(freerow, 1).Value = nest
                        End If
                        
                    End If
                    If InStr(1, nest, "epd", vbTextCompare) > 0 Then qVsht.Cells(freerow, 1).Value = nest
                    If InStr(1, nest, "cobie", vbTextCompare) > 0 Then qVsht.Cells(freerow, 1).Value = nest
                    If Len(qVsht.Cells(freerow, 1).Value) = 0 Then qVsht.Cells(freerow, 1).Value = "other"
                    If InStr(1, nest, "Data", vbTextCompare) > 0 Then
                        qVsht.Cells(freerow, 2).Value = nest
                     End If
                    
                    str = str + "  |  " + qVsht.Cells(freerow, freecol)
                    If InStr(1, nest, "ifc prop", vbTextCompare) > 0 Then
                            qVsht.Cells(freerow, freecol).Style = "Accent2"
                            qVsht.Cells(freerow, 4).Value = "mappable"
                            qVsht.Cells(freerow, 4).Style = "Explanatory Text"
                    Else: qVsht.Cells(freerow, freecol).Style = "Accent3"
                    End If
                    freecol = freecol + 1
                End If
                If Not InStr(1, qVsht.Cells(freerow, 2), "Data", vbTextCompare) > 0 Then
                         qVsht.Cells(freerow, 2).Value = "no category"
                         'qVsht.Cells(freerow, 2).Style = "Bad"
                End If
            Next j
            qVsht.Cells(freerow, 14) = str
            freecol = 8
            freerow = freerow + 1
'            For n = 1 To lastrowPset
'                If InStr(arr(n), prop) > 0 Then
'                   qVsht.Cells(freerow - 1, 1) = psht.Cells(n, 1)
'                    qVsht.Cells(freerow - 1, 1).Style = "Calculation"
'                    Exit For
'                End If
'            Next n
'                If Len(qVsht.Cells(freerow - 1, 1).Value) = 0 Then
'                    qVsht.Cells(freerow - 1, 1).Value = "IFC or other"
'                    qVsht.Cells(freerow - 1, 1).Style = "Linked Cell"
'                End If
            
            
        End If
    ElseIf batchsht.Cells(i, 1) = "Measure" Then
        If Len(batchsht.Cells(i, 3)) > 0 Then
            qVsht.Cells(freerow, 5).Style = "Input"
            qVsht.Cells(freerow, 5) = batchsht.Cells(i, 3)
            str = ""
            For j = 1 To propPsets.Count
                If batchsht.Cells(i, propPsets(j)) = "X" Then
                    qVsht.Cells(freerow, freecol) = batchsht.Cells(3, propPsets(j)) 'try with an X
                    str = str + "  |  " + qVsht.Cells(freerow, freecol)
                    qVsht.Cells(freerow, freecol).Style = "Accent6"
                Else: qVsht.Cells(freerow, freecol) = "X"
                      str = str + "  |  " + qVsht.Cells(freerow, freecol)
                      qVsht.Cells(freerow, freecol).Style = "Bad"
                End If
                freecol = freecol + 1
            Next j
            qVsht.Cells(freerow, 14) = str
            freecol = 8
            If Len(batchsht.Cells(i - 1, 3)) > 0 Then
                qVsht.Cells(freerow, 4) = batchsht.Cells(i - 1, 3)
                qVsht.Cells(freerow, 4).Style = "Input"
                Else: qVsht.Cells(freerow, 4).Style = "Bad"
            End If
            If Len(batchsht.Cells(i + 1, 3)) > 0 Then
                qVsht.Cells(freerow, 6) = batchsht.Cells(i + 1, 3)
                qVsht.Cells(freerow, 6).Style = "Input"
                Else: qVsht.Cells(freerow, 6).Style = "Bad"
             End If
            If Len(batchsht.Cells(i + 2, 3)) > 0 Then
                qVsht.Cells(freerow, 7) = batchsht.Cells(i + 2, 3)
                qVsht.Cells(freerow, 7).Style = "Input"
                Else: qVsht.Cells(freerow, 7).Style = "20% - Accent3"
            End If
                        
            freerow = freerow + 1
        End If
    
    End If

Next

lastrB = qVsht.Cells(qVsht.Rows.Count, "C").End(xlUp).Row
lastrD = qVsht.Cells(qVsht.Rows.Count, "E").End(xlUp).Row
If lastrB > lastrD Then
    lastr = lastrB
Else: lastr = lastrD
End If
temp = ""
For m = 2 To lastr
    If Len(qVsht.Cells(m, 1).Value) > 0 Then
        temp = qVsht.Cells(m, 1).Value
    Else: qVsht.Cells(m, 1).Value = temp
    End If
Next m
temp = ""
For m = 2 To lastr
    If Len(qVsht.Cells(m, 2).Value) > 0 Then
        temp = qVsht.Cells(m, 2).Value
    Else: qVsht.Cells(m, 2).Value = temp
    End If
                    If InStr(1, qVsht.Cells(m, 2).Value, "no category", vbTextCompare) > 0 Then
                        qVsht.Cells(m, 2).Style = "Bad"
                    End If
Next m



End Sub
