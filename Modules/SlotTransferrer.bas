
Public log As Logger

Dim fileName As String
Dim slotsSheet As String

Dim ownSlotsSheet As String

Dim dateCol As Integer
Dim dayCol As Integer
Dim armedCol As Integer
Dim armedStbCol As Integer
Dim unarmedCol As Integer
Dim unarmedStbCol As Integer
Dim guard2Col As Integer
Dim guard2StbCol As Integer
Dim pointsCol As Integer

Dim numDays As Integer

Dim plannerSheet As String
Dim depotCell As String
Dim fileNameCell As String

Dim depotName As String

' Variables to transfer S1 to 63

Dim depotAllocationCol As Integer
Dim guard2StartRow As Integer
Dim armedStartRow As Integer
Dim unarmedStartRow As Integer
Dim dutyDaysStartCol As Integer
Dim dutyDaysRow As Integer

Dim dutySlotsStartRow As Integer

Dim sourceSheet As Worksheet

Dim depotOffset As Integer

' Variables to transfer 63 to S1

Dim toFillSheet As String
Dim destinationSheet As Worksheet

Dim destiDateCol As Integer
Dim destiNumRowsPerDate As Integer
Dim destiDateStartRow As Integer

Dim destiGuard2DepotCol As Integer
Dim destiArmedDepotCol As Integer
Dim destiUnarmedDepotCol As Integer

Dim destiGuard2Col As Integer
Dim destiArmedCol As Integer
Dim destiUnarmedCol As Integer

Dim destiContactColOffset As Integer
Dim destiStandbyRowOffset As Integer

Dim destiNumSlotsPerDay As Integer

Sub TransferGuardSlots()
    ' initialises all variables
    initVars
    
    Set log = New Logger
    log.clearLog
    
    Dim numArmed As Integer
    Dim numUnarmed As Integer
    Dim num2IC As Integer
    
    Dim currentRow As Integer
    Dim currentDate As Integer
    
    Dim startRow As Integer
    Dim startCol As Integer
    Dim endRow As Integer
    Dim endCol As Integer
    
    Dim planningmonth As Date
    Dim currentDay As String
    
    ' Original shankar - Noted
    ' Dim destiCurrentRow As Integer
    Dim destiCurrentRow As Long
    destiCurrentRow = destiDateStartRow + 1
    
    Dim selectingRange As Range
    
    currentRow = dutySlotsStartRow
    currentDate = 1
    
    Dim i As Integer, j As Integer
    
    startRow = currentRow
    startCol = 1
    endRow = startRow
    endCol = pointsCol
    
    log.log ("Transferring from S1")
    
    ' Reset duty slots sheet
    With Worksheets(ownSlotsSheet)
        Set selectingRange = .Range(.Cells(startRow, startCol), .Cells(99, endCol))
    End With
    selectingRange.Borders.LineStyle = xlLineStyleNone
    selectingRange.Interior.ColorIndex = 0
    selectingRange.Value = ""
    
    planningmonth = DutySlots.getPlanningMonth
    
    i = 0
    While (i < numDays)
        
        'MsgBox (i & ": " & destiCurrentRow)
        
        ' Skip annoying blank days
        While (destinationSheet.Cells(destiCurrentRow, destiArmedDepotCol).Value = "")
            ' MsgBox ("Skip " & destiCurrentRow & " " & destiArmedDepotCol)
            Debug.Print destinationSheet.Cells(4, destiArmedDepotCol).Value
            destiCurrentRow = destiCurrentRow + 1
        Wend
    
        If (DutySlots.getDutyType = "GUARD") Then
        
            numArmed = 0
            numUnarmed = 0
            
            j = 0
            While (j < destiNumSlotsPerDay)
                If (destinationSheet.Cells(destiCurrentRow + j, destiArmedDepotCol).Value = depotName) Then
                    numArmed = numArmed + 1
                End If
                If (destinationSheet.Cells(destiCurrentRow + j, destiUnarmedDepotCol).Value = depotName) Then
                    numUnarmed = numUnarmed + 1
                End If
                j = j + 1
            Wend
            
            ' Create 1 row for each slot, minimum 1 for when no duty on that day
            j = 0
            While (j = 0 Or j < numArmed Or j < numUnarmed)
            
                currentDay = UCase(Format(currentDate & "/" & month(planningmonth) & "/" & Year(planningmonth), "ddd"))
            
                ' Initialise row
                Worksheets(ownSlotsSheet).Cells(currentRow, dateCol).Value = currentDate
                Worksheets(ownSlotsSheet).Cells(currentRow, dateCol).HorizontalAlignment = xlCenter
                Worksheets(ownSlotsSheet).Cells(currentRow, dayCol).Value = currentDay
                Worksheets(ownSlotsSheet).Cells(currentRow, dayCol).HorizontalAlignment = xlCenter
                
                ' Set white cell if there is a slot
                If (j < numArmed) Then
                    If (currentDay = "SAT" Or currentDay = "SUN") Then
                        Worksheets(ownSlotsSheet).Cells(currentRow, armedCol).Interior.ColorIndex = 15
                        Worksheets(ownSlotsSheet).Cells(currentRow, armedStbCol).Interior.ColorIndex = 15
                    Else
                        Worksheets(ownSlotsSheet).Cells(currentRow, armedCol).Interior.ColorIndex = 0
                        Worksheets(ownSlotsSheet).Cells(currentRow, armedStbCol).Interior.ColorIndex = 0
                    End If
                Else
                    Worksheets(ownSlotsSheet).Cells(currentRow, armedCol).Interior.ColorIndex = 1
                    Worksheets(ownSlotsSheet).Cells(currentRow, armedStbCol).Interior.ColorIndex = 1
                End If
                If (j < numUnarmed) Then
                    If (currentDay = "SAT" Or currentDay = "SUN") Then
                        Worksheets(ownSlotsSheet).Cells(currentRow, unarmedCol).Interior.ColorIndex = 15
                        Worksheets(ownSlotsSheet).Cells(currentRow, unarmedStbCol).Interior.ColorIndex = 15
                    Else
                        Worksheets(ownSlotsSheet).Cells(currentRow, unarmedCol).Interior.ColorIndex = 0
                        Worksheets(ownSlotsSheet).Cells(currentRow, unarmedStbCol).Interior.ColorIndex = 0
                    End If
                Else
                    Worksheets(ownSlotsSheet).Cells(currentRow, unarmedCol).Interior.ColorIndex = 1
                    Worksheets(ownSlotsSheet).Cells(currentRow, unarmedStbCol).Interior.ColorIndex = 1
                End If
                
                If (currentDay = "SAT" Or currentDay = "SUN") Then
                    Worksheets(ownSlotsSheet).Cells(currentRow, pointsCol).Value = 2
                    
                    ' Set grey for weekends
                    Worksheets(ownSlotsSheet).Cells(currentRow, dateCol).Interior.ColorIndex = 15
                    Worksheets(ownSlotsSheet).Cells(currentRow, dayCol).Interior.ColorIndex = 15
                    Worksheets(ownSlotsSheet).Cells(currentRow, pointsCol).Interior.ColorIndex = 15
                Else
                    Worksheets(ownSlotsSheet).Cells(currentRow, pointsCol).Value = 1
                End If
                Worksheets(ownSlotsSheet).Cells(currentRow, armedCol).HorizontalAlignment = xlCenter
                Worksheets(ownSlotsSheet).Cells(currentRow, armedStbCol).HorizontalAlignment = xlCenter
                Worksheets(ownSlotsSheet).Cells(currentRow, unarmedCol).HorizontalAlignment = xlCenter
                Worksheets(ownSlotsSheet).Cells(currentRow, unarmedStbCol).HorizontalAlignment = xlCenter
                Worksheets(ownSlotsSheet).Cells(currentRow, armedCol).VerticalAlignment = xlCenter
                Worksheets(ownSlotsSheet).Cells(currentRow, armedStbCol).VerticalAlignment = xlCenter
                Worksheets(ownSlotsSheet).Cells(currentRow, unarmedCol).VerticalAlignment = xlCenter
                Worksheets(ownSlotsSheet).Cells(currentRow, unarmedStbCol).VerticalAlignment = xlCenter
                Worksheets(ownSlotsSheet).Cells(currentRow, armedCol).Font.Bold = False
                Worksheets(ownSlotsSheet).Cells(currentRow, armedStbCol).Font.Bold = False
                Worksheets(ownSlotsSheet).Cells(currentRow, unarmedCol).Font.Bold = False
                Worksheets(ownSlotsSheet).Cells(currentRow, unarmedStbCol).Font.Bold = False
                Worksheets(ownSlotsSheet).Cells(currentRow, pointsCol).HorizontalAlignment = xlCenter
                
                endRow = currentRow
                currentRow = currentRow + 1
                j = j + 1
            Wend
        ElseIf (DutySlots.getDutyType = "GUARD 2IC") Then
        
            currentDay = UCase(Format(currentDate & "/" & month(planningmonth) & "/" & Year(planningmonth), "ddd"))
            
            ' Initialise row
            Worksheets(ownSlotsSheet).Cells(currentRow, dateCol).Value = currentDate
            Worksheets(ownSlotsSheet).Cells(currentRow, dateCol).HorizontalAlignment = xlCenter
            Worksheets(ownSlotsSheet).Cells(currentRow, dayCol).Value = currentDay
            Worksheets(ownSlotsSheet).Cells(currentRow, dayCol).HorizontalAlignment = xlCenter
            
            ' Set white cell if there is a slot
            If (destinationSheet.Cells(destiCurrentRow, destiGuard2DepotCol).Value = depotName) Then
                If (currentDay = "SAT" Or currentDay = "SUN") Then
                    Worksheets(ownSlotsSheet).Cells(currentRow, guard2Col).Interior.ColorIndex = 15
                    Worksheets(ownSlotsSheet).Cells(currentRow, guard2StbCol).Interior.ColorIndex = 15
                    
                Else
                    Worksheets(ownSlotsSheet).Cells(currentRow, guard2Col).Interior.ColorIndex = 0
                    Worksheets(ownSlotsSheet).Cells(currentRow, guard2StbCol).Interior.ColorIndex = 0
                End If
            Else
                Worksheets(ownSlotsSheet).Cells(currentRow, guard2Col).Interior.ColorIndex = 1
                Worksheets(ownSlotsSheet).Cells(currentRow, guard2StbCol).Interior.ColorIndex = 1
            End If
            
            If (currentDay = "SAT" Or currentDay = "SUN") Then
                Worksheets(ownSlotsSheet).Cells(currentRow, pointsCol).Value = 2
                Worksheets(ownSlotsSheet).Cells(currentRow, pointsCol).Interior.ColorIndex = 15
                
                ' Set grey for weekends
                Worksheets(ownSlotsSheet).Cells(currentRow, dateCol).Interior.ColorIndex = 15
                Worksheets(ownSlotsSheet).Cells(currentRow, dayCol).Interior.ColorIndex = 15
            Else
                Worksheets(ownSlotsSheet).Cells(currentRow, pointsCol).Value = 1
            End If
            Worksheets(ownSlotsSheet).Cells(currentRow, guard2Col).HorizontalAlignment = xlCenter
            Worksheets(ownSlotsSheet).Cells(currentRow, guard2StbCol).HorizontalAlignment = xlCenter
            Worksheets(ownSlotsSheet).Cells(currentRow, pointsCol).HorizontalAlignment = xlCenter
            
            endRow = currentRow
            currentRow = currentRow + 1
        End If
        
        currentDate = currentDate + 1
        i = i + 1
        
        ' Find start of next day
        destiCurrentRow = destiCurrentRow - 1
        destiCurrentRow = destiCurrentRow + destiNumRowsPerDate
        destiCurrentRow = destiCurrentRow + 1
        
    Wend
    
    
    With Worksheets(ownSlotsSheet)
        Set selectingRange = .Range(.Cells(startRow, startCol), .Cells(endRow, endCol))
    End With
    selectingRange.Borders.LineStyle = xlContinuous
    selectingRange.Font.ColorIndex = 1
    selectingRange.Font.Size = 8
    selectingRange.Font.name = "Calibri"
    
    With Worksheets(ownSlotsSheet)
        Set selectingRange = .Range(.Cells(startRow, dateCol), .Cells(endRow, dayCol))
    End With
    selectingRange.Font.Size = 11
    selectingRange.Font.name = "Calibri"
    
    With Worksheets(ownSlotsSheet)
        Set selectingRange = .Range(.Cells(startRow, dayCol), .Cells(endRow, dayCol))
    End With
    
    With Worksheets(ownSlotsSheet)
        Set selectingRange = .Range(.Cells(startRow, pointsCol), .Cells(endRow, pointsCol))
    End With
    selectingRange.Font.Size = 11
    selectingRange.Font.name = "Calibri"
    
    log.log ("Transfer complete")
    
End Sub

Sub TransferToS1()
    initVars
    
    Set log = New Logger
    log.clearLog
    
    Dim currentRow As Integer
    Dim currentDate As Integer
    
    Dim guardName As String
    Dim standbyName As String
    
    Dim lastDateProcessed As Integer
    lastDateProcessed = 0
    Dim lastDestiRowProcessed As Integer
    lastDestiRowProcessed = destiDateStartRow
    Dim numSkip As Integer
    
    Dim destiCurrentRow As Integer
    
    Dim foundSlot As Boolean
    
    Dim i As Integer, j As Integer
    
    log.log ("Transferring to S1")
    
    ' Go through each planned duty
    currentRow = 3
    While (Worksheets(ownSlotsSheet).Cells(currentRow, dateCol).Value <> "")
        currentDate = Worksheets(ownSlotsSheet).Cells(currentRow, dateCol).Value
        ' When there is more than 1 duty on the same day, skip over last transferred
        If (currentDate = lastDateProcessed) Then
            numSkip = numSkip + 1
        Else
            numSkip = 0
            lastDateProcessed = currentDate
        End If
        
        ' Find correct day
        i = lastDestiRowProcessed
        While (destinationSheet.Cells(i, destiDateCol).Value <> currentDate)
            i = i + destiNumRowsPerDate
        Wend
        destiCurrentRow = i + 1
        
        ' Skip annoying blank days
        While (destinationSheet.Cells(destiCurrentRow, destiArmedDepotCol).Value = "")
            destiCurrentRow = destiCurrentRow + 1
        Wend
        
        lastDestiRowProcessed = destiCurrentRow - 1
        
        If (DutySlots.getDutyType = "GUARD") Then
            
            ' Transfer armed
            If (Worksheets(ownSlotsSheet).Cells(currentRow, armedCol).Interior.Color <> RGB(0, 0, 0)) Then
            
                ' Warn empty slots
                If (Worksheets(ownSlotsSheet).Cells(currentRow, armedCol).Value = "") Then
                    MsgBox ("WARNING!!! There are empty slots!")
                ElseIf (Worksheets(ownSlotsSheet).Cells(currentRow, armedStbCol).Value = "") Then
                    MsgBox ("WARNING!!! There are empty slots!")
                End If
            
                guardName = Worksheets(ownSlotsSheet).Cells(currentRow, armedCol).Value
                standbyName = Worksheets(ownSlotsSheet).Cells(currentRow, armedStbCol).Value
                foundSlot = False
                i = 0
                j = 0
                While (i < destiNumSlotsPerDay)
                    
                    If (destinationSheet.Cells(destiCurrentRow + i, destiArmedDepotCol).Value = depotName) Then
                        If (j = numSkip And foundSlot = False) Then
                            foundSlot = True
                            destinationSheet.Cells(destiCurrentRow + i, destiArmedCol).Value = PointsTable.getRank(guardName) & " " & guardName
                            destinationSheet.Cells(destiCurrentRow + i, destiArmedCol + destiContactColOffset).Value = PointsTable.getContact(guardName)
                            
                            destinationSheet.Cells(destiCurrentRow + destiStandbyRowOffset + i, destiArmedCol).Value = PointsTable.getRank(standbyName) & " " & standbyName
                            destinationSheet.Cells(destiCurrentRow + destiStandbyRowOffset + i, destiArmedCol + destiContactColOffset).Value = PointsTable.getContact(standbyName)
                        Else
                            j = j + 1
                        End If
                    End If
                    i = i + 1
                Wend
            End If
            
            ' Transfer unarmed
            If (Worksheets(ownSlotsSheet).Cells(currentRow, unarmedCol).Interior.Color <> RGB(0, 0, 0)) Then
            
                ' Warn empty slots
                If (Worksheets(ownSlotsSheet).Cells(currentRow, unarmedCol).Value = "") Then
                    MsgBox ("WARNING!!! There are empty slots!")
                ElseIf (Worksheets(ownSlotsSheet).Cells(currentRow, unarmedStbCol).Value = "") Then
                    MsgBox ("WARNING!!! There are empty slots!")
                End If
                
                guardName = Worksheets(ownSlotsSheet).Cells(currentRow, unarmedCol).Value
                standbyName = Worksheets(ownSlotsSheet).Cells(currentRow, unarmedStbCol).Value
                foundSlot = False
                i = 0
                j = 0
                While (i < destiNumSlotsPerDay)
                    If (destinationSheet.Cells(destiCurrentRow + i, destiUnarmedDepotCol).Value = depotName) Then
                        If (j = numSkip) Then
                            foundSlot = True
                            destinationSheet.Cells(destiCurrentRow + i, destiUnarmedCol).Value = PointsTable.getRank(guardName) & " " & guardName
                            destinationSheet.Cells(destiCurrentRow + i, destiUnarmedCol + destiContactColOffset).Value = PointsTable.getContact(guardName)
                            
                            destinationSheet.Cells(destiCurrentRow + destiStandbyRowOffset + i, destiUnarmedCol).Value = PointsTable.getRank(standbyName) & " " & standbyName
                            destinationSheet.Cells(destiCurrentRow + destiStandbyRowOffset + i, destiUnarmedCol + destiContactColOffset).Value = PointsTable.getContact(standbyName)
                        Else
                            j = j + 1
                        End If
                    End If
                    i = i + 1
                Wend
            End If
        ElseIf (DutySlots.getDutyType = "GUARD 2IC") Then
            
            ' Transfer guard 2
            If (Worksheets(ownSlotsSheet).Cells(currentRow, guard2Col).Interior.Color <> RGB(0, 0, 0)) Then
                guardName = Worksheets(ownSlotsSheet).Cells(currentRow, guard2Col).Value
                standbyName = Worksheets(ownSlotsSheet).Cells(currentRow, guard2StbCol).Value
                foundSlot = False
                i = 0
                j = 0
                While (i < destiNumSlotsPerDay)
                    If (destinationSheet.Cells(destiCurrentRow + i, destiGuard2DepotCol).Value = depotName) Then
                        If (j = numSkip) Then
                            foundSlot = True
                            destinationSheet.Cells(destiCurrentRow + i, destiGuard2Col).Value = PointsTable.getRank(guardName) & " " & guardName
                            destinationSheet.Cells(destiCurrentRow + i, destiGuard2Col + destiContactColOffset).Value = PointsTable.getContact(guardName)
                            
                            destinationSheet.Cells(destiCurrentRow + destiStandbyRowOffset + i, destiGuard2Col).Value = PointsTable.getRank(standbyName) & " " & standbyName
                            destinationSheet.Cells(destiCurrentRow + destiStandbyRowOffset + i, destiGuard2Col + destiContactColOffset).Value = PointsTable.getContact(standbyName)
                        Else
                            j = j + 1
                        End If
                    End If
                    i = i + 1
                Wend
            End If
        End If
        
        currentRow = currentRow + 1
    Wend
    
    log.log ("Transfer complete")
End Sub

Private Sub initVars()
    plannerSheet = "Guard Duty Planner"
    depotCell = "M2"
    fileNameCell = "M3"
    
    ownSlotsSheet = "Duty Slots"
    
    dateCol = 1
    dayCol = 2
    armedCol = 3
    armedStbCol = 4
    unarmedCol = 5
    unarmedStbCol = 6
    guard2Col = 3
    guard2StbCol = 4
    
    Dim i As Integer
    i = 3
    While (Worksheets(ownSlotsSheet).Cells(2, i).Value <> "POINTS")
        numDutyCols = numDutyCols + 1
        i = i + 2
    Wend
    pointsCol = i
    
    fileName = Worksheets(plannerSheet).Range(fileNameCell).Value
    slotsSheet = "Forecast"
    
    depotName = Worksheets(plannerSheet).Range(depotCell).Value
    
    depotAllocationCol = 1
    guard2StartRow = 13
    armedStartRow = 21
    unarmedStartRow = 29
    
    dutyDaysStartCol = 2
    dutyDaysRow = 3
    
    dutySlotsStartRow = 3
    
    Dim curb As Workbook
    Set curb = ActiveWorkbook
    Workbooks.Open (ActiveWorkbook.Path & "\" & fileName)
    curb.Activate
    Set sourceSheet = Workbooks(fileName).Worksheets(slotsSheet)
    
    i = 0
    While (sourceSheet.Cells(guard2StartRow + i, depotAllocationCol).Value <> depotName And i < 6)
        i = i + 1
    Wend
    
    If (i = 6) Then
        MsgBox ("Unable to find depot. Did you enter the correct name?")
        depotOffset = -1
    Else
        depotOffset = i
    End If
    
    Dim planningmonth As Date
    planningmonth = DutySlots.getPlanningMonth
    
    If (month(planningmonth) = 1 Or month(planningmonth) = 3 Or month(planningmonth) = 5 Or month(planningmonth) = 7 Or month(planningmonth) = 8 Or month(planningmonth) = 10 Or month(planningmonth) = 12) Then
        numDays = 31
    ElseIf (month(planningmonth) = 4 Or month(planningmonth) = 6 Or month(planningmonth) = 9 Or month(planningmonth) = 11) Then
        numDays = 30
    ElseIf ((Year(planningmonth) Mod 4 = 0 And Year(planningmonth) Mod 100 <> 0) Or (Year(planningmonth) Mod 400 = 0)) Then
        numDays = 29
    Else
        numDays = 28
    End If
    
    ' Transfer 63 to S1
    
    toFillSheet = "Allocation"
    Set destinationSheet = Workbooks(fileName).Worksheets(toFillSheet)

    destiDateCol = 1
    destiNumRowsPerDate = 9
    destiDateStartRow = 5
    
    destiGuard2DepotCol = 3
    destiArmedDepotCol = 6
    destiUnarmedDepotCol = 9
    
    destiGuard2Col = 2
    destiArmedCol = 5
    destiUnarmedCol = 8
    
    destiContactColOffset = 2
    destiStandbyRowOffset = 4
    
    destiNumSlotsPerDay = 3
End Sub
