

Sub prepare_report()
'
' prepare_report Macro
'

With Application
.ScreenUpdating = False
.Calculation = xlCalculationManual

End With

' Set the width of the progress bar to 0.

 
Dim StartTime As Double
Dim SecondsElapsed As Double
Dim pasted As Worksheet
Dim types As Worksheet
Dim clasif As Worksheet
Dim parents As Worksheet
Dim lastRow As Long
Dim TargetRange As Range
Dim TargetRange2 As Range
Dim FoundCell As Range
Dim FoundCell2 As Range
Dim i As Long
Dim j As Long
Dim PctDone As Single

'Remember time when macro starts
  StartTime = Timer


Set pasted = ActiveWorkbook.Worksheets("paste report")
Set types = ActiveWorkbook.Worksheets("types")
Set clasif = ActiveWorkbook.Worksheets("classifications")
Set TargetRange = ActiveWorkbook.Worksheets("classifications").Range("C2:S63441")
Set TargetRange2 = ActiveWorkbook.Worksheets("Parents-Children").Range("C2:C2000")
Set parents = ActiveWorkbook.Worksheets("Parents-Children")

If pasted.Range("Z1").Value = 1 Then
    MsgBox ("Report from Products has already been prepared")
    Exit Sub
End If

pasted.Rows(1).EntireRow.Delete

lastRow = pasted.Cells(pasted.Rows.Count, "C").End(xlUp).row + 1

pasted.Range("B1:B" & lastRow).UnMerge

For i = 1 To lastRow
    If pasted.Cells(i, 2) > 0 Then
        pasted.Cells(i, 6) = pasted.Cells(i, 2)
    Else
        pasted.Cells(i, 6) = pasted.Cells(i - 1, 6)
    End If
Next


For i = 1 To lastRow
        'index so it doesnt go to do smth for eternity
        j = 1
        'rows
        types.Cells(i, 1) = i
        'types
        types.Cells(i, 3) = pasted.Cells(i, 6)
        'classification
        types.Cells(i, 6) = pasted.Cells(i, 4)
        'Codes
        types.Cells(i, 4) = pasted.Cells(i, 5)
        'search for definitions of code
        On Error Resume Next
        Do While j < 10
            Set FoundCell = TargetRange.Find(types.Cells(i, 4))
            
            If clasif.Cells(FoundCell.row, 2) = types.Cells(i, 6) Then Exit Do
             
            j = j + 1
        Loop
        types.Cells(i, 5).Value = clasif.Cells(FoundCell.row, 19)
        
        'search for definitions of code
        On Error Resume Next
        Set FoundCell2 = TargetRange2.Find(types.Cells(i, 3))
        types.Cells(i, 2).Value = parents.Cells(FoundCell2.row, 6).Value
        
        Application.StatusBar = "Progress: " & i & " of " & lastRow & " " & Format(i / lastRow, "0%")
        
        'If i Mod 100 = 0 Then Debug.Print i
        
'        ' Update the percentage completed.
'        PctDone = i / lastRow
'
'        ' Call subroutine that updates the progress bar.
'        UpdateProgressBar PctDone
Next


With Application
.ScreenUpdating = True
.Calculation = xlCalculationAutomatic

End With


'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'insert value 1 in cell Z2 so macro can check if the prepare report code has already been done
pasted.Range("Z1").Value = 1


'Notify user in seconds
  MsgBox "Classification Report was succesfully prepared and loaded in the Excel workbook in " & SecondsElapsed & " seconds", vbInformation

    
End Sub


