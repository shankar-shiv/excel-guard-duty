Dim extrasSheet As String
Dim startRow As Integer
Dim nameCol As Integer
Dim monthCol As Integer

' Count number of months someone is exempted
Function CountTotalMonthExtras() As Integer
    
    initVars
    
    Dim numRows As Integer
    numRows = countRows

    Dim currentRow
    Dim i As Integer
    i = 0
    
    Dim extraMonths As Integer
    
    While (i < numRows)
        currentRow = startRow + i
        
        If (DutySlots.getPlanningMonth = Worksheets(extrasSheet).Cells(currentRow, monthCol).Value) Then
            extraMonths = extraMonths + 1
        End If
        
        i = i + 1
    Wend
    
    CountTotalMonthExtras = extraMonths
End Function

Function PersonnelNumExtras(name As String) As Integer
    initVars
    
    Dim numRows As Integer
    numRows = countRows

    Dim currentRow
    Dim i As Integer
    i = 0
    
    Dim numExtra As Integer
    numExtra = 0
    
    While (i < numRows)
        currentRow = startRow + i
        
        If (name = Worksheets(extrasSheet).Cells(currentRow, nameCol).Value) Then
            If (DutySlots.getPlanningMonth = Worksheets(extrasSheet).Cells(currentRow, monthCol).Value) Then
                numExtra = numExtra + 1
            End If
        End If
        
        i = i + 1
    Wend
    
    PersonnelNumExtras = numExtra
End Function

Private Sub initVars()
    extrasSheet = "Extra Duties"
    startRow = 2
    nameCol = 1
    monthCol = 2
End Sub

Function countRows()
    Dim i As Integer
    
    i = startRow
    While (Not IsEmpty(Worksheets(extrasSheet).Cells(i, 1).Value))
        i = i + 1
    Wend
    
    countRows = i - startRow
End Function


