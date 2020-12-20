Dim exemptionsSheet As String
Dim startRow As Integer
Dim nameCol As Integer
Dim monthCol As Integer

' Count number of months someone is exempted
Function CountExemptionMonths(name As String) As Integer
    
    initVars
    
    Dim numRows As Integer
    numRows = countRows

    Dim currentRow
    Dim i As Integer
    i = 0
    
    Dim exemptionMonths As Integer
    
    While (i < numRows)
        currentRow = startRow + i
        
        If (name = Worksheets(exemptionsSheet).Cells(currentRow, nameCol).Value And DateDiff("m", DutySlots.getPlanningMonth, Worksheets(exemptionsSheet).Cells(currentRow, monthCol).Value) <= 0) Then
            exemptionMonths = exemptionMonths + 1
        End If
        
        i = i + 1
    Wend
    
    CountExemptionMonths = exemptionMonths
End Function

Function PersonnelHasExemption(name As String) As Boolean
    initVars
    
    Dim numRows As Integer
    numRows = countRows ' numRow = 0 becaus there are no names present

    Dim currentRow
    Dim i As Integer
    i = 0
    
    Dim exemptionMonths As Integer
    
    While (i < numRows)
        currentRow = startRow + i
        
        If (name = Worksheets(exemptionsSheet).Cells(currentRow, nameCol).Value) Then
            If (DutySlots.getPlanningMonth = Worksheets(exemptionsSheet).Cells(currentRow, monthCol).Value) Then
                PersonnelHasExemption = True
            End If
        End If
        
        i = i + 1
    Wend
End Function

Private Sub initVars()
    exemptionsSheet = "Duty Exemptions"
    startRow = 2
    nameCol = 1
    monthCol = 2
End Sub

Function countRows()
    Dim i As Integer
    
    i = startRow
    While (Not IsEmpty(Worksheets(exemptionsSheet).Cells(i, 1).Value))
        i = i + 1
    Wend
    
    countRows = i - startRow
End Function
