

Function GetSearchArrayTypesD(strSearch)
Dim strResults As String
Dim SHT As Worksheet
Dim rFND As Range
Dim sFirstAddress
Set SHT = ThisWorkbook.Worksheets(2)
    Set rFND = Nothing
    With SHT.Range("D2:D100000")
        Set rFND = .Cells.Find(What:=strSearch, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlRows, SearchDirection:=xlNext, MatchCase:=False)
        If Not rFND Is Nothing Then
            sFirstAddress = rFND.Address
            Do
                If strResults = vbNullString Then
                    strResults = rFND.Address
                Else
                    strResults = strResults & "|" & rFND.Address
                End If
                Set rFND = .FindNext(rFND)
            Loop While Not rFND Is Nothing And rFND.Address <> sFirstAddress
        End If
    End With

If strResults = vbNullString Then
    GetSearchArray = Null
ElseIf InStr(1, strResults, "|", 1) = 0 Then
    GetSearchArray = Array(strResults)
Else
    GetSearchArray = Split(strResults, "|")
End If
End Function

'Sub test2()
'For Each X In GetSearchArray("screw")
 '   Debug.Print X
'Next
'End Sub

Function GetSearchArrayClassifS(strSearch)
Dim strResults As String
Dim SHT As Worksheet
Dim rFND As Range
Dim sFirstAddress
Set SHT = ThisWorkbook.Worksheets(1)
    Set rFND = Nothing
    With SHT.Range("S2:S70000")
        Set rFND = .Cells.Find(What:=strSearch, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlRows, SearchDirection:=xlNext, MatchCase:=False)
        If Not rFND Is Nothing Then
            sFirstAddress = rFND.Address
            Do
                If strResults = vbNullString Then
                    strResults = rFND.Address
                Else
                    strResults = strResults & "|" & rFND.Address
                End If
                Set rFND = .FindNext(rFND)
            Loop While Not rFND Is Nothing And rFND.Address <> sFirstAddress
        End If
    End With

If strResults = vbNullString Then
    GetSearchArray = Null
ElseIf InStr(1, strResults, "|", 1) = 0 Then
    GetSearchArray = Array(strResults)
Else
    GetSearchArray = Split(strResults, "|")
End If
End Function

Sub find_strings_2()

Dim ArrayCh() As Variant
Dim C As Range
Dim firstAddress As String
Dim i As Integer

 ArrayCh = Array("a", "b", "c") 'strings to lookup

With ActiveSheet.Cells
    For i = LBound(ArrayCh) To UBound(ArrayCh)
        Set C = .Find(What:=ArrayCh(i), LookAt:=xlPart, LookIn:=xlValues)

        If Not C Is Nothing Then
            firstAddress = C.Address 'used later to verify if looping over the same address
            Do
               
                Debug.Print ArrayCh(i) & " " & C.Address 'example
                '_____
                Set C = .FindNext(C)

            Loop While Not C Is Nothing And C.Address <> firstAddress
        End If
    Next i
End With

End Sub
