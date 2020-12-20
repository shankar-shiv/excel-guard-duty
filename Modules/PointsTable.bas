
Dim pointsTableSheet As String

Dim startRow As Integer
Dim rankCol As Integer
Dim nameCol As Integer
Dim contactCol As Integer
Dim dutyTypeCol As Integer
Dim ppmCol As Integer

' Calculate points per month
Sub CalculatePPM()

    initVars
    
    Dim name As String
    
    Dim numRows As Integer
    numRows = countRows
    
    Dim currentRow
    Dim i As Integer
    i = 0
    
    While (i < numRows)
        currentRow = startRow + i
        
        name = Worksheets(pointsTableSheet).Cells(currentRow, nameCol).Value
        Worksheets(pointsTableSheet).Cells(currentRow, ppmCol).Value = DutyRecords.averagePoints(name)
        
        i = i + 1
    Wend
    
End Sub

Private Sub initVars()
    pointsTableSheet = "Points Table"
    startRow = 2
    rankCol = 2
    nameCol = 3
    contactCol = 4
    dutyTypeCol = 7
    ppmCol = 9
End Sub

Function countRows()
    initVars
    
    Dim i As Integer
    
    i = startRow
    While (Not IsEmpty(Worksheets(pointsTableSheet).Cells(i, 1).Value))
        i = i + 1
    Wend
    
    countRows = i - startRow
End Function

Function getRank(name As String)
    initVars
    
    i = startRow
    While (Not IsEmpty(Worksheets(pointsTableSheet).Cells(i, 1).Value))
        If (Worksheets(pointsTableSheet).Cells(i, nameCol).Value = name) Then
            getRank = Worksheets(pointsTableSheet).Cells(i, rankCol).Value
        End If
        i = i + 1
    Wend
End Function

Function getName(rowNum As String) ' rowNum is obtained from Points Sheet ex: Row 2 to 42.
    initVars
    
    i = startRow ' startRow = 2
                                                 ' .Cells(2, 1)
    While (Not IsEmpty(Worksheets(pointsTableSheet).Cells(i, 1).Value))
        If (i = rowNum) Then
            getName = Worksheets(pointsTableSheet).Cells(i, nameCol).Value
        End If
        i = i + 1
    Wend
End Function

Function getContact(name As String)
    initVars
    
    i = startRow
    While (Not IsEmpty(Worksheets(pointsTableSheet).Cells(i, 1).Value))
        If (Worksheets(pointsTableSheet).Cells(i, nameCol).Value = name) Then
            getContact = Worksheets(pointsTableSheet).Cells(i, contactCol).Value
        End If
        i = i + 1
    Wend
End Function

Function getDutyType(rowNum As String)
    initVars
    
    i = startRow
    While (Not IsEmpty(Worksheets(pointsTableSheet).Cells(i, 1).Value))
        If (i = rowNum) Then
            getDutyType = Worksheets(pointsTableSheet).Cells(i, dutyTypeCol).Value
        End If
        i = i + 1
    Wend
End Function
