



Private Sub choose1_Click()
'Me.ListBox4.Clear

Dim row As Long
    Dim i As Long
    Dim col As Integer

row = Me.ListBox1.Column(3, Me.ListBox1.ListIndex)
i = ListBox4.ListCount

                 Me.ListBox4.AddItem
                 Me.ListBox4.List(i, 0) = Worksheets(1).Cells(row, 2).Value
                 Me.ListBox4.List(i, 1) = Worksheets(1).Cells(row, 3).Value
                 Me.ListBox4.List(i, 2) = Worksheets(1).Cells(row, 19).Value
                 Me.ListBox4.List(i, 3) = row
       
End Sub

Private Sub choose2_Click()

'Me.ListBox4.Clear

Dim row As Long
    Dim i As Long
    Dim col As Integer

row = Me.ListBox3.Column(3, Me.ListBox3.ListIndex)
i = ListBox4.ListCount

                Me.ListBox4.AddItem
                Me.ListBox4.List(i, 3) = row
                Me.ListBox4.List(i, 2) = Worksheets(1).Cells(row, 19).Value
                Me.ListBox4.List(i, 1) = Worksheets(1).Cells(row, 3).Value
                Me.ListBox4.List(i, 0) = Worksheets(1).Cells(row, 2).Value
                
End Sub

Private Sub cmb_remove_from_list_Click()

'Dim itm As Variant
Dim srem As String
Dim asrem As Variant

'    For Each itm In ListBox4.ItemsSelected
'        srem = srem & "," & itm
'    Next
'
'    asrem = Split(Mid(srem, 2), ",")
'    For i = UBound(asrem) To 0 Step -1
'        ListBox4.RemoveItem ListBox4.ItemData(asrem(i))
'    Next

If ListBox4.ListIndex >= 0 Then
         ListBox4.RemoveItem ListBox4.ListIndex
    End If

End Sub

Private Sub cmb11_Click()

Me.ListBox2.Clear

Dim searchStr As String
 Dim C As Range
    Dim i As Long

searchStr = Me.ListBox1.Column(1, Me.ListBox1.ListIndex)
            
            With Worksheets("types").Range("D2:D70000")
            Set C = .Find(What:=searchStr, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox2.AddItem
                                        
                Me.ListBox2.List(i, 2) = Worksheets("types").Cells(C.row, 5).Value
                Me.ListBox2.List(i, 1) = Worksheets("types").Cells(C.row, 3).Value
                Me.ListBox2.List(i, 0) = Worksheets("types").Cells(C.row, 4).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox2.AddItem "Not found"
        End If
        End With



End Sub



Private Sub cmb2_Click()

   
     Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'ListBox1.Clear
   
   Me.ListBox1.Clear
    Me.ListBox2.Clear
   
   Dim startrow As Long
   Dim lastRow As Long
   Dim database() As String
       
    Dim C As Range
    Dim i As Long
    Dim j As Long
    Dim FindThis As String
    Dim Str
    Dim results() As String
    Dim entry As Classification_Entry
    Dim pos As Integer
     Dim strKey As Variant
       
       ' Select Case ComboBox1.Value
              
'        'Case "Uniclass 1.4":
'            If Me.tb1.Value = True Then
'                 FindThis = SearchBox.Value
'                 'With Worksheets(1).Range("S2:S1965")
'                 For Each strKey In dict
'                    entry = dict(strKey)
'                    pos = InStr(entry.Description, FindThis)
'                    If Not pos = 0 Then
'                             Me.ListBox1.AddItem
'                             Me.ListBox1.List(i, 2) = entry.Description
'                             Me.ListBox1.List(i, 1) = entry.Code
'                             Me.ListBox1.List(i, 0) = entry.Clasification
'                             i = i + 1
'
'                 Else: Me.ListBox1.AddItem "Not found"
'                 End If
'                 Next
'            End If
'
          

            
          'Case "Uniclass 1.4":
            If Me.tb1.Value = True Then
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S2:S1965")
            
            
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
            Else: Me.ListBox1.AddItem "Not found"
            End If
            End With
            End If
            
            
            
         'Case "Uniclass 2":
           If Me.tb2.Value = True Then
           FindThis = SearchBox.Value
            With Worksheets(1).Range("S1966:S8742")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
         'Case "BSAB":
         If Me.tb4.Value = True Then
           
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S24313:S36585")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
            
            ' Case "Uniclass 2015":
            If Me.tb3.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S8743:S17406")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
             'Case "UNSPSC":
             If Me.tb6.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S46340:S61179")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
            
             'Case "TFM":
             If Me.tb11.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S61244:S62059")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
             
             
             'Case "revit":
             If Me.tb10.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S61180:S61243")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
             'Case "Omniclass 12":
             If Me.tb9.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S17407:S24312")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
             'Case "NRM3":
             If Me.tb7.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S36706:S37045")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
            ' Case "CPV":
            If Me.tb5.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S37046:S46339")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
        'Case "CI/SfB":
        If Me.tb8.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S36586:S36705")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        'Case "SFG":
         If Me.tb12.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S62060:S62624")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
         'Case "All":
         If Me.tb13.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S2:S63441")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
        'Case "NS 3451":
         If Me.tb14.Value = True Then
            
            FindThis = SearchBox.Value
            With Worksheets(1).Range("S62625:S63441")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = C.Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
       ' End Select
    
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub UserForm1_Initialize()

Application.WindowState = xlMinimized
Me.ListBox1.Clear
Me.ListBox2.Clear
Me.ListBox3.Clear
Me.ListBox4.Clear
Me.SearchBox.Value = ""
Me.tb99.Value = ""
Me.lbRLO.Clear
Me.lbTFM.Clear
Me.lbIFC.Clear





End Sub






Private Sub cmb22_Click()

Me.ListBox2.Clear

Dim searchStr As String
 Dim C As Range
    Dim i As Long

searchStr = Me.ListBox3.Column(1, Me.ListBox3.ListIndex)
            
            With Worksheets("types").Range("D2:D70000")
            Set C = .Find(What:=searchStr, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox2.AddItem
                                        
                Me.ListBox2.List(i, 2) = Worksheets("types").Cells(C.row, 5).Value
                Me.ListBox2.List(i, 1) = Worksheets("types").Cells(C.row, 3).Value
                Me.ListBox2.List(i, 0) = Worksheets("types").Cells(C.row, 4).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox2.AddItem "Not found"
        End If
        End With


End Sub

Private Sub cmbexport_Click()
Dim r As Integer
Dim C As Integer
Dim lastRow As Integer

Set SHT = Worksheets("export")

lastRow = SHT.Cells(SHT.Rows.Count, "A").End(xlUp).row + 1

For r = 0 To ListBox4.ListCount - 1
    For C = 0 To 3
        Worksheets("export").Range("A" & lastRow).Offset(r, C).Value = ListBox4.List(r, C)
    Next
Next


End Sub

Private Sub cmbinfo_Click()

Me.tb99.Value = ""
Me.lbRLO.Clear
Me.lbTFM.Clear
Me.lbIFC.Clear


Dim row As Long
    Dim i As Long
    Dim col As Integer

row = Me.ListBox1.Column(3, Me.ListBox1.ListIndex)


                 
            If Worksheets(1).Cells(row, 30).Value > 0 Then
                Me.tb99.Value = Worksheets(1).Cells(row, 30).Value
             
             End If
            
         
            
        For col = 31 To 36
           
                Me.lbRLO.AddItem Worksheets(1).Cells(row, col).Value
                
             Next col
             
             
      For col = 37 To 41
           
                Me.lbTFM.AddItem Worksheets(1).Cells(row, col).Value
                
             Next col
             
               
      For col = 42 To 45
           
                Me.lbIFC.AddItem Worksheets(1).Cells(row, col).Value
                
             Next col
             
             
             
End Sub

Private Sub CommandButton26_Click()

Me.ListBox1.Clear
Me.ListBox2.Clear
Me.ListBox3.Clear
Me.ListBox4.Clear
Me.SearchBox.Value = ""
Me.tb99.Value = ""
Me.lbRLO.Clear
Me.lbTFM.Clear
Me.lbIFC.Clear


End Sub

Private Sub CommandButton27_Click()

Dim row As Long
    Dim i As Long
    Dim col As Integer

row = Me.ListBox1.Column(3, Me.ListBox1.ListIndex)

ActiveWorkbook.Worksheets("classifications").Activate
ActiveWorkbook.Worksheets("classifications").Cells(row, 2).Select
    
    
End Sub



Private Sub CommandButton28_Click()

   
     Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'ListBox1.Clear
   
   Me.ListBox1.Clear
    Me.ListBox2.Clear
   
   Dim startrow As Long
   Dim lastRow As Long
   Dim database() As String
       
    Dim C As Range
    Dim i As Long
    Dim j As Long
    Dim FindThis As String
    Dim Str
    Dim results() As String
    Dim entry As Classification_Entry
    Dim pos As Integer
     Dim strKey As Variant
       
       ' Select Case ComboBox1.Value
              
'        'Case "Uniclass 1.4":
'            If Me.tb1.Value = True Then
'                 FindThis = SearchBox.Value
'                 'With Worksheets(1).Range("S2:S1965")
'                 For Each strKey In dict
'                    entry = dict(strKey)
'                    pos = InStr(entry.Description, FindThis)
'                    If Not pos = 0 Then
'                             Me.ListBox1.AddItem
'                             Me.ListBox1.List(i, 2) = entry.Description
'                             Me.ListBox1.List(i, 1) = entry.Code
'                             Me.ListBox1.List(i, 0) = entry.Clasification
'                             i = i + 1
'
'                 Else: Me.ListBox1.AddItem "Not found"
'                 End If
'                 Next
'            End If
'
          

            
          'Case "Uniclass 1.4":
            If Me.tb1.Value = True Then
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C2:C1965")
            
            
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
            Else: Me.ListBox1.AddItem "Not found"
            End If
            End With
            End If
            
            
            
         'Case "Uniclass 2":
           If Me.tb2.Value = True Then
           FindThis = SearchBox1.Value
            With Worksheets(1).Range("C1966:C8742")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
         'Case "BSAB":
         If Me.tb4.Value = True Then
           
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C24313:C36585")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
            
            ' Case "Uniclass 2015":
            If Me.tb3.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C8743:C17406")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
             'Case "UNSPSC":
             If Me.tb6.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C46340:C61179")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
            
             'Case "TFM":
             If Me.tb11.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C61244:C62059")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
             
             
             'Case "revit":
             If Me.tb10.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C61180:C61243")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
             'Case "Omniclass 12":
             If Me.tb9.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C17407:C24312")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
             'Case "NRM3":
             If Me.tb7.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C36706:C37045")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
            ' Case "CPV":
            If Me.tb5.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C37046:C46339")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
        'Case "CI/SfB":
        If Me.tb8.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C36586:C36705")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        'Case "SFG":
         If Me.tb12.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C62060:C62624")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
         'Case "All":
         If Me.tb13.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C2:C63441")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
        'Case "NS 3451":
         If Me.tb14.Value = True Then
            
            FindThis = SearchBox1.Value
            With Worksheets(1).Range("C62625:C63441")
            Set C = .Find(What:=FindThis, LookAt:=xlPart, LookIn:=xlValues)
            i = 0
            If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
                Me.ListBox1.AddItem
                Me.ListBox1.List(i, 3) = Worksheets(1).Cells(C.row, 1).Value
                Me.ListBox1.List(i, 2) = Worksheets(1).Cells(C.row, 19).Value
                Me.ListBox1.List(i, 1) = Worksheets(1).Cells(C.row, 3).Value
                Me.ListBox1.List(i, 0) = Worksheets(1).Cells(C.row, 2).Value
                i = i + 1
                Set C = .FindNext(C)
            Loop While Not C Is Nothing And C.Address <> firstAddress
        Else: Me.ListBox1.AddItem "Not found"
        End If
        End With
        End If
        
        
       ' End Select
    
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Me.ListBox3.Clear

Dim row As Long
    Dim i As Long
    Dim col As Integer

row = Me.ListBox1.Column(3, Me.ListBox1.ListIndex)

i = 0
        For col = 26 To 21 Step -1
           
            If Worksheets(1).Cells(row, col).Value > 0 Then
                Me.ListBox3.AddItem
                Me.ListBox3.List(i, 3) = row
                Me.ListBox3.List(i, 2) = Worksheets(1).Cells(Cells(row, col).Value, 19)
                Me.ListBox3.List(i, 1) = Worksheets(1).Cells(Cells(row, col).Value, 3)
                Me.ListBox3.List(i, 0) = Worksheets(1).Cells(Cells(row, col).Value, 2)
                i = i + 1
             Else:
             End If
             Next col
End Sub



Private Sub tb99_Change()

End Sub


Private Sub UserForm_Click()

End Sub



Private Sub UserForm_Terminate()

Set dict = Nothing

End Sub
