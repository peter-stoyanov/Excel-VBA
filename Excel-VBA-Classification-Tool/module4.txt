Sub LevelEntries()

Dim row As Long
Dim col As Long

For row = 2 To 62059
'For row = 48971 To 61243

    For col = 4 To 10
    
        If Len(Cells(row, col).Value) = 0 Then
                
            Cells(row, 20).Value = col - 4
            Exit For
        
        End If
    
    
    Next col
    
    If Not Len(Cells(row, 10).Value) = 0 Then
                
            Cells(row, 20).Value = 7
                    
        End If
    
Next row

MsgBox ("Ready")



End Sub


Sub FindParents()

Dim row As Long
Dim flag As Long
Dim colOffset As Long
Dim parentrow1 As Long
Dim parentrow2 As Long
Dim parentrow3 As Long
Dim parentrow4 As Long
Dim parentrow5 As Long
Dim parentrow6 As Long
Dim parentrow7 As Long

iterator = 0
colOffset = 1

For row = 49032 To 49032
'For row = 48971 To 61243

    
If Cells(row, 20).Value = 1 Then GoTo NextIteration
  
           
        Do
        iterator = iterator + 1
        parentrow1 = row - iterator
        If Not (Cells(row - iterator, 20).Value = Cells(row, 20).Value Or Cells(row - iterator, 20).Value >= Cells(row, 20).Value) And Cells(row - iterator, 20).Value >= 1 Then
                        
                Cells(row, 20 + colOffset).Value = row - iterator
                 colOffset = colOffset + 1
                
            End If
        Loop While (Cells(row - iterator, 20).Value = Cells(row, 20).Value Or Cells(row - iterator, 20).Value >= Cells(row, 20).Value) And Cells(row - iterator, 20).Value >= 1
        
        
        
        Do
         iterator = iterator + 1
        parentrow2 = row - iterator + 1
        If Not (Cells(row - iterator, 20).Value = Cells(parentrow1, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow1, 20).Value) And Cells(row - iterator, 20).Value >= 1 Then
                        
                Cells(row, 20 + colOffset).Value = row - iterator
                 colOffset = colOffset + 1
                
            End If
        Loop While (Cells(row - iterator, 20).Value = Cells(parentrow1, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow2, 20).Value) And Cells(row - iterator, 20).Value >= 1
        
        
         
         
         Do
          iterator = iterator + 1
        parentrow3 = row - iterator + 1
        If Not (Cells(row - iterator, 20).Value = Cells(parentrow2, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow2, 20).Value) And Cells(row - iterator, 20).Value >= 1 Then
                        
                Cells(row, 20 + colOffset).Value = row - iterator
                 colOffset = colOffset + 1
                
            End If
        Loop While (Cells(row - iterator, 20).Value = Cells(parentrow2, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow2, 20).Value) And Cells(row - iterator, 20).Value >= 1
  
  
            
            
            Do
             iterator = iterator + 1
        parentrow4 = row - iterator + 1
        If Not (Cells(row - iterator, 20).Value = Cells(parentrow3, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow3, 20).Value) And Cells(row - iterator, 20).Value >= 1 Then
                        
                Cells(row, 20 + colOffset).Value = row - iterator
                 colOffset = colOffset + 1
                
            End If
            Loop While (Cells(row - iterator, 20).Value = Cells(parentrow3, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow3, 20).Value) And Cells(row - iterator, 20).Value >= 1
  
            
            
            Do
            iterator = iterator + 1
        parentrow5 = row - iterator + 1
        If Not (Cells(row - iterator, 20).Value = Cells(parentrow4, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow4, 20).Value) And Cells(row - iterator, 20).Value >= 1 Then
                        
                Cells(row, 20 + colOffset).Value = row - iterator
                 colOffset = colOffset + 1
                
            End If
            Loop While (Cells(row - iterator, 20).Value = Cells(parentrow4, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow4, 20).Value) And Cells(row - iterator, 20).Value >= 1
      
            
            
            Do
            iterator = iterator + 1
        parentrow6 = row - iterator + 1
        If Not (Cells(row - iterator, 20).Value = Cells(parentrow5, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow5, 20).Value) And Cells(row - iterator, 20).Value >= 1 Then
                        
                Cells(row, 20 + colOffset).Value = row - iterator
                 colOffset = colOffset + 1
                
            End If
            Loop While (Cells(row - iterator, 20).Value = Cells(parentrow5, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow5, 20).Value) And Cells(row - iterator, 20).Value >= 1
      
            
            
            Do
             iterator = iterator + 1
        parentrow7 = row - iterator + 1
        If Not (Cells(row - iterator, 20).Value = Cells(parentrow6, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow6, 20).Value) And Cells(row - iterator, 20).Value >= 1 Then
                        
                Cells(row, 20 + colOffset).Value = row - iterator
                 colOffset = colOffset + 1
                
            End If
            Loop While (Cells(row - iterator, 20).Value = Cells(parentrow6, 20).Value Or Cells(row - iterator, 20).Value >= Cells(parentrow6, 20).Value) And Cells(row - iterator, 20).Value >= 1
      
      
NextIteration:
iterator = 0
colOffset = 1

Next row


MsgBox ("Ready")


End Sub


Sub FindParents2()

Application.ScreenUpdating = False
Application.Calculation = xlManual

Dim row As Long
Dim searchStr As Long
Dim firstAddress As String
Dim cRow As Long

'36585


    For row = 24313 To 36585
    searchStr = Worksheets(1).Cells(row, 20).Value - 2
        If searchStr >= 1 Then
                 
            cRow = FindBack(searchStr, row)
            Worksheets(1).Cells(row, 22).Value = cRow
        End If
        
        If row Mod 1000 = 0 Then
        Debug.Print row
        End If
        
    Next row
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

MsgBox "ready"

End Sub

Function FindBack(x As Long, row As Long)

Dim i As Long
FindBack = 0
    
    For i = 1 To row
        If Worksheets(1).Cells(row - i, 20).Value = x Then
            FindBack = row - i
            Exit For
        End If
    Next i

End Function
