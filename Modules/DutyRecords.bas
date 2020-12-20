
' This Module points to Duty Record tab

' All the basic information defined.
Dim recordsSheet As String
Dim startRow As Integer
Dim nameCol As Integer
Dim monthCol As Integer
Dim dutyTypeCol As Integer
Dim pointsCol As Integer

' Add duty record for a person as a new row
Sub AddDutyRecord(name As String, month As String, dutyType As String, points As Integer)
    Call initVars
    
    Dim numRows As Integer
    numRows = countRows()
    
    Dim addRow As Integer
    addRow = startRow + numRows
    
    Worksheets(recordsSheet).Cells(addRow, nameCol).Value = name
    Worksheets(recordsSheet).Cells(addRow, monthCol).Value = month
    Worksheets(recordsSheet).Cells(addRow, dutyTypeCol).Value = dutyType
    Worksheets(recordsSheet).Cells(addRow, pointsCol).Value = points
End Sub

Sub ReverseDutyRecord()
    initVars
    
    Dim planningmonth As Date
    planningmonth = DutySlots.getPlanningMonth
    
    Dim numRows As Integer
    numRows = countRows
    
    Dim i As Integer
    
    i = 0
    While (i < numRows)
        If (Worksheets(recordsSheet).Cells(startRow + i, monthCol).Value = planningmonth) Then
            'Worksheets(recordsSheet).Cells(startRow + i, nameCol).Value = ""
            'Worksheets(recordsSheet).Cells(startRow + i, monthCol).Value = ""
            'Worksheets(recordsSheet).Cells(startRow + i, dutyTypeCol).Value = ""
            'Worksheets(recordsSheet).Cells(startRow + i, pointsCol).Value = ""
            Worksheets(recordsSheet).Rows(startRow + i).EntireRow.Delete
        Else
            i = i + 1
        End If
    Wend
End Sub

' Average out points that a person has done based on their duty record
Function averagePoints(name As String) As Double
    
    ' Call the initialisation function
    Call initVars
    
    'Calculate the total number of rows.
    Dim numRows As Integer
    numRows = countRows() ' numRows = 41
    ' Debug.Print numRows
    
    
    Dim points As Double
    points = 0
    
    Dim startMonth As Date
    startMonth = "01/01/2001"
    
    Dim lastMonth As Date
    lastMonth = "01/01/2001"
    
    Dim currentRow
    Dim i As Integer
    i = 0
    
    ' Add total points
    While (i < numRows) ' i < 41
        currentRow = startRow + i ' 2 + 0...until 40
        
        If (name = Worksheets(recordsSheet).Cells(currentRow, nameCol).Value) Then
            If (startMonth = "01/01/2001") Then
                startMonth = Worksheets(recordsSheet).Cells(currentRow, monthCol).Value
                Debug.Print "startMonth : ", startMonth
            End If
            points = points + Worksheets(recordsSheet).Cells(currentRow, pointsCol).Value
        End If
        
        If (lastMonth < Worksheets(recordsSheet).Cells(currentRow, monthCol).Value) Then
            lastMonth = Worksheets(recordsSheet).Cells(currentRow, monthCol).Value
        End If
        
        i = i + 1
    Wend
    
    points = points - DutyExtras.PersonnelNumExtras(name) * 2
    
    Dim numMonths As Integer
    numMonths = DateDiff("m", startMonth, lastMonth) + 1
    numMonths = numMonths - DutyExemptions.CountExemptionMonths(name)
    
    If (numMonths = 0) Then
        averagePoints = 0
    Else
        averagePoints = points / numMonths
    End If
End Function

Private Sub initVars()
    recordsSheet = "Duty Record"
    startRow = 2
    nameCol = 1
    monthCol = 2
    dutyTypeCol = 3
    pointsCol = 4
End Sub

Function countRows()
    Dim i As Integer
    
    i = startRow
    While (Not IsEmpty(Worksheets(recordsSheet).Cells(i, 1).Value))
        i = i + 1
    Wend
    
    countRows = i - startRow
End Function

