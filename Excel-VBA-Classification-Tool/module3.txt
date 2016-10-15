Sub FlatList()

Dim row As Long
Dim col As Long

For row = 62060 To 62624

    For col = 11 To 18
    
        If Not Cells(row, col).Value = 0 Then
                
            Cells(row, 19).Value = Cells(row, col).Value
        
        
        End If
    
    
    Next col
    
Next row

End Sub
