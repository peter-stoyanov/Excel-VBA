Sub flow()

With Application
.ScreenUpdating = False
.Calculation = xlCalculationManual
.DisplayStatusBar = False

End With

Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
  StartTime = Timer


Application.Run "prepare_old_data"
Application.Run "prepare_new_data"

Application.Run "concat_old_data"
Application.Run "concat_new_data"
'Application.Run "createSheet"
Application.Run "comparesheets"



With Application
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic
.DisplayStatusBar = True
.Calculate
End With

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)


'Notify user in seconds
  MsgBox "Comparison was succesfully prepared and loaded in " & SecondsElapsed & " seconds", vbInformation


End Sub


Sub prepare_old_data()
'
' prepare old sheet data = unmerge and fill type on all rows
'

Dim sht As Worksheet
Dim lastRow As Long


Set sht = ActiveWorkbook.Worksheets("old")

lastRow = sht.Cells(sht.Rows.Count, "C").End(xlUp).Row + 1

sht.Range("B1:B" & lastRow).UnMerge

For i = 3 To lastRow
    If sht.Cells(i, 2) > 0 Then
        sht.Cells(i, 6) = sht.Cells(i, 2)
    Else
        sht.Cells(i, 6) = sht.Cells(i - 1, 6)
    End If
Next

   
End Sub

Sub prepare_new_data()
'
' prepare old sheet data = unmerge and fill type on all rows
'

Dim sht As Worksheet
Dim lastRow As Long


Set sht = ActiveWorkbook.Worksheets("new")

lastRow = sht.Cells(sht.Rows.Count, "C").End(xlUp).Row + 1

sht.Range("B1:B" & lastRow).UnMerge

For i = 3 To lastRow
    If sht.Cells(i, 2) > 0 Then
        sht.Cells(i, 6) = sht.Cells(i, 2)
    Else
        sht.Cells(i, 6) = sht.Cells(i - 1, 6)
    End If
Next

   
End Sub


Sub concat_old_data()
'
' prepare old sheet data = concatenate Type + Class + Code
'

Dim sht As Worksheet
Dim lastRow As Long


Set sht = ActiveWorkbook.Worksheets("old")

lastRow = sht.Cells(sht.Rows.Count, "F").End(xlUp).Row + 1

For i = 3 To lastRow
    sht.Cells(i, 15) = sht.Cells(i, 6) & "  |  " & sht.Cells(i, 4) & "  |  " & sht.Cells(i, 5)
Next

   
End Sub

Sub concat_new_data()
'
' prepare new sheet data = concatenate Type + Class + Code
'

Dim sht As Worksheet
Dim lastRow As Long


Set sht = ActiveWorkbook.Worksheets("new")

lastRow = sht.Cells(sht.Rows.Count, "F").End(xlUp).Row + 1

For i = 3 To lastRow
    sht.Cells(i, 15) = sht.Cells(i, 6) & "  |  " & sht.Cells(i, 4) & "  |  " & sht.Cells(i, 5)
Next

   
End Sub

Sub comparesheets()
'
' compare old with new
'

Dim old_sht As Worksheet
Dim new_sht As Worksheet
Dim lastRow_new As Long
Dim lastRow_old As Long

Dim i As Long
Dim arrSum As Variant, arrUsers As Variant
Dim cUnique As New Collection

'Put the name range from "Summary" in an array
With ThisWorkbook.Sheets("new")
    arrSum = .Range("O3", .Range("O" & Rows.Count).End(xlUp))
End With

'"Convert" the array to a collection (unique items)
For i = 1 To UBound(arrSum, 1)
    On Error Resume Next
    cUnique.Add arrSum(i, 1), CStr(arrSum(i, 1))
Next i

'Get the users array
With ThisWorkbook.Sheets("old")
    arrUsers = .Range("O3", .Range("O" & Rows.Count).End(xlUp))
End With

ActiveWorkbook.Worksheets("whats new").Select
Selection.NumberFormat = "@"

Dim counter As Long
Dim strTest As String
Dim strArray() As String
Dim intCount As Integer
counter = 1
'Check if the value exists in the Users sheet
For i = 1 To cUnique.Count
    'if can't find the value in the users range, delete the rows
    If Application.WorksheetFunction.VLookup(cUnique(i), arrUsers, 1, False) = "#N/A" Then
       strTest = cUnique(i)
       strArray = Split(strTest, "  |  ")
       ThisWorkbook.Sheets("whats new").Cells(counter, 2) = Trim(strArray(0))
       ThisWorkbook.Sheets("whats new").Cells(counter, 3) = Trim(strArray(1))
       ThisWorkbook.Sheets("whats new").Cells(counter, 4) = Trim(strArray(2))
       counter = counter + 1
    End If
Next i
'removes AutoFilter if one remains
ThisWorkbook.Sheets("Summary").AutoFilterMode = False
  
    ActiveWorkbook.Worksheets("whats new").Activate
    Columns("B:D").Select
    Selection.Columns.AutoFit
    Columns("D:D").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1").Select

   
End Sub
