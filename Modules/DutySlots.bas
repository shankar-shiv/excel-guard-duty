
Public log As Logger

' Variables for initVars sub
Dim dutySlotsSheet As String
Dim dutySlotsStartRow As Integer

Dim dateCol As Integer
Dim dayCol As Integer
Dim firstActualCol As Integer
Dim firstStbCol As Integer
Dim numDutyCols As Integer
Dim pointsCol As Integer

Dim firstCheckCol As Integer
Dim lastCheckCol As Integer

Dim dutyTypeCell As String
Dim monthCell As String
Dim dutyHeaderRow As Integer

Dim dutyType As String
Dim planningmonth As Date
' end initVars Sub

Dim slotPoints(2) As Integer ' slotPoints(numberOfPoints) = number of slots with that number of points
Dim slotPoint As Integer

Dim personnel(100) As DutyPersonnel ' Original : Dim personnel(255) as DutyPersonnel
Dim numPersonnel As Integer

Dim slots(100) As DutySlot ' Original : Dim slots(255) As DutySlot
Dim numSlots As Integer

' Variables for Guard Duty planner sheet defined in initVars sub
Dim plannerSheet As String
Dim dutyGapCell As String
Dim standbyGapCell As String

Dim minDutyGap As Integer
Dim minStbGap As Integer
' end

' Automate planning of duties
Sub PlanDuties()
    
    initVars ' Initialise variables
    
    Dim numRows As Integer
    Dim currentRow As Integer
    
    Dim totalPoints As Integer
    totalPoints = 0
    Dim numDuties As Integer
    numDuties = 0
    Dim overUnder As Double
    
    ' Track unfulfilled duties to swap
    
    
    Dim cont As Boolean
    
    Dim ds As DutySlot ' Create ds object from the DutySlot class module
    
    ' Iterative variables
    Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer
    
    Dim slotPointsBackup(2) As Integer
    
    numPersonnel = 0
    numSlots = 0
    
    slotPoints(0) = 0
    slotPoints(1) = 0
    slotPoints(2) = 0
    
    numRows = countRows ' numRows = 52
    
    Set log = New Logger ' Create new Logger object
    log.clearLog ' Clear screen
    'log.log ("Number of duties per day: " & numDutyCols)
    
    Dim dayHasDuty As Boolean
    
    ' Count total points, number of slots, load all slots
    i = 0
    While (i < numRows) ' i < 52
        currentRow = dutySlotsStartRow + i ' currentRow = 3 + i(0)
        
        dayHasDuty = False ' If cell is not black
        j = 0
        While (j < numDutyCols) ' j < 2
                                               '(3        ,       3 + 2 * 0       )
            If (Worksheets(dutySlotsSheet).Cells(currentRow, firstActualCol + 2 * j).Interior.Color <> RGB(0, 0, 0)) Then
                dayHasDuty = True
            End If
            j = j + 1
        Wend
        
        If (dayHasDuty) Then
            j = 0
            While (j < numDutyCols) ' j < 2
                                                   '(4         ,        3 + 2 * 0      )
                If (Worksheets(dutySlotsSheet).Cells(currentRow, firstActualCol + 2 * j).Interior.Color <> RGB(0, 0, 0)) Then ' Unarmed
                    ' If the day has duty aka not in black colour, the slotPoint will be added up
                    ' slotPoint can only be either 1 or 2.
                    slotPoint = Worksheets(dutySlotsSheet).Cells(currentRow, pointsCol).Value
                    totalPoints = totalPoints + slotPoint
                    numDuties = numDuties + 1
                    slotPoints(slotPoint) = slotPoints(slotPoint) + 1
                    
                    ' Create an object of DutySlot with every slot available (not black in colour)
                    Set slots(numSlots) = New DutySlot
                    
                    ' + Initialise the slot by calling the initialize function
                    '   from DutySlot Class Module
                    Call slots(numSlots).initialize(currentRow, firstActualCol + 2 * j)
                    
                    ' Increment the next slot
                    numSlots = numSlots + 1
                End If
                j = j + 1
            Wend
            
        End If
        
        i = i + 1
    Wend
    
    ' Create backup for slotPoints, not sure why ?
    i = 0
    While (i < 3)
        slotPointsBackup(i) = slotPoints(i)
        i = i + 1
    Wend
    
    overUnder = totalPoints / numDuties
    
    log.log ("Total Points: " & totalPoints)
    log.log ("Total Duties: " & numDuties)
    log.log ("Avg Points Per Slot: " & overUnder)
    
    ' Load all the guards from Points Sheet
    Dim n As Integer
    n = PointsTable.countRows ' n = 41
    i = 0
    
    Dim dp As DutyPersonnel
    While (i < n)
        
        ' Don't count guards who have exemptions                                                                                                            ' dutyType = GUARD
        If (Not DutyExemptions.PersonnelHasExemption(PointsTable.getName(i + 2)) And PointsTable.getDutyType(i + 2) = dutyType) Then
            
            ' Create new guards objects from DutyPersonnel Class Module
            Set personnel(numPersonnel) = New DutyPersonnel
            personnel(numPersonnel).initialize (PointsTable.getName(i + 2))
            
            numPersonnel = numPersonnel + 1
        End If
        i = i + 1
    Wend
    
    log.log ("Number of Duty Personnel: " & numPersonnel) ' numPersonnel = 41
    log.log ("Points Per Person: " & (totalPoints / numPersonnel) & vbCrLf) ' vbCrLf means press Enter
    
    ' Sort personnel by points in a linear fashion.
    sortPersonnelByPoints
    
    ' Remove extras from assigning list
    slotPoints(2) = slotPoints(2) - DutyExtras.CountTotalMonthExtras
    ' Debug.Print slotPoints(0) = 0
    ' Debug.Print slotPoints(1) = 40
    ' Debug.Print slotPoints(2) = 12
    If (slotPoints(2) < 0) Then slotPoints(2) = 0
    
    ' Setup slots
    Dim currentDay As Integer
    Dim numVolunteers As Integer
    numVolunteers = 0
    
    ' This whole While ... Wend chunk of code is for
    ' 1) Calculating the difficulty of a slot to be assigned to guards
    ' 2) Allocating volunteer duties
    
    ' Note
    ' i is for iterating slots; j is for iterating guards personnel
    i = 0
    While (i < numSlots) ' While (i < 52 free slots where guards will be assigned)
        currentDay = slots(i).day ' TODO : This is NOT currentDay, it is current Date.
        
        j = 0
        While (j < numPersonnel) ' j < 41
            ' If personnel A-AO has committment on the 2nd Then add 1 to the slots(0).difficulty variable.
                                                                ' This repeat for all the 52 slots.
            ' If personnel A-AO has committment on the 2nd Then add 1 to the slots(1).difficulty variable.
            ' If personnel A-AO has committment on the 2nd Then add 1 to the slots(2).difficulty variable.
            
            If (personnel(j).getCommitment(currentDay)) Then ' Set difficulty
                slots(i).difficulty = slots(i).difficulty + 1
            End If
            
            ' Set volunteer
            If (personnel(j).getVolunteer(currentDay) And Not personnel(j).getDutyDay(currentDay)) Then
                slots(i).setVolunteer (personnel(j).name)
                personnel(j).addDutyDay (currentDay)
            End If
            j = j + 1
        Wend
        
        If (slots(i).personnel <> "") Then ' Pre-allocated / Volunteer personnel
            j = 0
            While (j < numPersonnel) ' j < 41
                If (personnel(j).name = slots(i).personnel) Then
                    slots(i).setVolunteer (personnel(j).name)
                    
                    slotPoints(slots(i).points) = slotPoints(slots(i).points) - 1
                    numVolunteers = numVolunteers + 1
                    
                    If (personnel(j).getVolunteer(slots(i).day)) Then
                        log.log (personnel(j).name & " volunteered on " & currentDay)
                    Else
                        log.log (personnel(j).name & " was pre-assigned on " & currentDay)
                    End If
                    ' The below 2 lines are subtract and add up the points for that day..
                    ' Debug.Print slotPoints(0) = 0
                    ' Debug.Print slotPoints(1) = 40
                    ' Debug.Print slotPoints(2) = 12
                    personnel(j).removeDuty (slots(i).points)
                    personnel(j).addDutyDay (slots(i).day)
                    slots(i).locked = True
                End If
                j = j + 1
            Wend
        End If
        
        i = i + 1
    Wend
    
    ' Assign number of duties
    ' Assign guards with 2 point slots. If the total number of 2 point slots are empty
    ' then Assign guards with 1 point slots.
    Dim cpp As Integer ' Current Planning Points (points of the slot)
    Dim currIndex As Integer
    Dim breakPoint As Integer
    breakPoint = -1
    cpp = 2
    currIndex = numPersonnel - 1 ' 40 = 41 - 1
    i = 0
    
            ' (i < 52 - 0 - 49)
    While (i < numDuties - DutyExtras.CountTotalMonthExtras - numVolunteers)
        If (cpp = 2) Then
            personnel(currIndex).addDuty (cpp) ' personnel(40).addDuty(2)
            slotPoints(cpp) = slotPoints(cpp) - 1 ' slotPoints(2) = slotPoints(2) - 1
        ElseIf (cpp = 1) Then
            'If (personnel(currIndex).numberOfDutiesWithPoints(2) > 0 Or slotPoints(cpp) = 1) Then
                personnel(currIndex).addDuty (cpp)
                slotPoints(cpp) = slotPoints(cpp) - 1
            'Else
            '    personnel(currIndex).addDuty (cpp)
            '    personnel(currIndex).addDuty (cpp)
            '    slotPoints(cpp) = slotPoints(cpp) - 2
            '    i = i + 1
            'End If
        End If
        
        If (slotPoints(cpp) = 0) Then ' If (slotPoints(2) = 0) Then
            cpp = cpp - 1             ' after all the '2 point' slots are assigned to guards
                                      ' move on to '1 point' slots.
            If (cpp = 1 And currIndex > 0) Then breakPoint = currIndex - 1
        End If
        
        currIndex = currIndex - 1 ' 39 = 40 -1
        If (currIndex = -1) Then
            If (breakPoint = -1) Then
                currIndex = numPersonnel - 1
            Else
                currIndex = breakPoint
                If (personnel(currIndex).numberOfDutiesWithPoints(1) >= 2) Then
                    currIndex = numPersonnel - 1
                End If
            End If
        End If
        
        i = i + 1
    Wend
    
    ' Balance duties based on points
    cont = True
    currIndex = numPersonnel - 1 ' 40 = 41 - 1
    i = 0
    
    'Debug.Print "personnel(40).totalPoints", personnel(currIndex).totalPoints ' = 2
    'Debug.Print "personnel(0).totalPoints", personnel(i).totalPoints          ' = 0
    'Debug.Print "personnel(0).dutyPoints", personnel(i).dutyPoints            ' = 10
    'Debug.Print "personnel(40).dutyPoints", personnel(currIndex).dutyPoints   ' = 0.182
    
    While (cont)
        If (personnel(currIndex).totalPoints = personnel(i).totalPoints And personnel(i).dutyPoints > 0 And personnel(currIndex).dutyPoints > 0) Then
                '(10 - 0.182) / 10 -- What's this formula ?
                ' The personnel who has more dutyPoints will be swapped with personnel
                ' who has less dutyPoints if this condition is met (personnel(i).numberOfDutiesWithPoints(1) > 0)
            If ((personnel(i).dutyPoints - personnel(currIndex).dutyPoints) / personnel(i).dutyPoints > 0.1) Then
                If (personnel(i).numberOfDutiesWithPoints(1) > 0) Then
                    personnel(i).removeDuty (1) ' Swap duty with personnel who has less dutyPoints
                    personnel(currIndex).addDuty (1)
                End If
            End If
        Else
            cont = False
        End If
        ' Compare the guards from 0,1,2,3 and the guards from 40,39,38,37.
        ' Compare guards 0 And 40 , 1 And 39 , 2 And 38 ...
        currIndex = currIndex - 1
        i = i + 1
        
    Wend
    
    ' Assign extras
    i = 0
    Dim numExtras As Integer
    
    While (i < numPersonnel) 'While (i < 41)
        j = 0
        numExtras = DutyExtras.PersonnelNumExtras(personnel(i).name)
        While (j < numExtras)
            personnel(i).addDuty (2)
            log.log (personnel(i).name & " has extra")
            j = j + 1
        Wend
        
        i = i + 1
    Wend
    
    i = 0
    While (i < 3)
        slotPoints(i) = slotPointsBackup(i)
        i = i + 1
    Wend
    
    ' Assign standbys
    sortPersonnelByPointsReverse
    
    cpp = 2
    currIndex = numPersonnel - 1
    i = 0
    While (i < numDuties)
        If (cpp = 2) Then
            personnel(currIndex).addStandby (cpp)
            slotPoints(cpp) = slotPoints(cpp) - 1
        ElseIf (cpp = 1) Then
            If (personnel(currIndex).numberOfDutiesWithPoints(2) > 0 Or slotPoints(cpp) = 1) Then
                personnel(currIndex).addStandby (cpp)
                slotPoints(cpp) = slotPoints(cpp) - 1
            Else
                personnel(currIndex).addStandby (cpp)
                personnel(currIndex).addStandby (cpp)
                slotPoints(cpp) = slotPoints(cpp) - 2
                i = i + 1
            End If
        End If
        
        If (slotPoints(cpp) = 0) Then
            cpp = cpp - 1
        End If
        
        currIndex = currIndex - 1
        If (currIndex = -1) Then currIndex = numPersonnel - 1
        
        i = i + 1
    Wend
    
    ' This subroutine sorts all 52 slots in a decreasing order of number of points.
    sortSlotsByPoints
    
    ' Sort slots and personnel by difficulty.
    ' The personnel who has many committments will have a higher number of pDifficulty value
    ' than a personnel who has little committments.
    sortPersonnelByDifficulty
    
    ' Print "" to the cell.
    log.log ("")
    
    
    ' Assign duties
    
    Dim numDutiesToAssign() As Integer
    Dim foundSlot As Boolean
    Dim hasClash As Boolean
    Dim hasMissingSlot As Boolean
    hasMissingSlot = True
    
    Dim numRetries As Integer
    numRetries = 0
    
    Dim errorLog(255) As String
    Dim numErrors As Integer
    numErrors = 0
    
  ' While (True         And          0 < 1)
    While (hasMissingSlot And numRetries < 1)
        hasMissingSlot = False
        numRetries = numRetries + 1
        numErrors = 0
        
        i = 0
        While (i < numSlots) ' While (i < 52 free slots where guards will be assigned)
            If (Not slots(i).locked) Then
                slots(i).personnel = "" 'Assign "" the slot which is not assigned to any personnel
                slots(i).standby = ""
            End If
            i = i + 1 'Loop i from 0 to 51 aka loop 52 times
        Wend
        
        i = 0
        While (i < numPersonnel) ' Loop through each of the 41 total personnel
            j = 1
            
            ' Initially numDutiesToAssign is 0.
            numDutiesToAssign = personnel(i).numberOfDuties
            
            'numDutiesToAssign consists of the below 3 variables
            'numDuties(0) = 0
            'numDuties(1) = 0
            'numDuties(2) = 0
            
            ' TODO : I cannot understand the value from the below code.
            ' Debug.Print "numDutiesToAssign", numDutiesToAssign(2)
            
            ' (1 <= 2)
            While (j <= 2) ' Loop through each duty points 1 => Weekday, 2 => Weekend
                k = 0
                
                ' While ( 0 ... < numDutiesToAssign(1) or numDutiesToAssign(2))
                While (k < numDutiesToAssign(j)) ' Loop through each duty
                    l = 0
                    foundSlot = False
                    ' (0 ... 51 < 52)
                    While (l < numSlots) ' Loop through each 52 slots.
                        
                        ' Checks if there is a slot open or it has found solution
                        If (Not foundSlot And slots(l).personnel = "" And slots(l).points = j) Then
                            m = 0
                            hasClash = False
                            
                            ' (0 ... 51 < 52)
                            ' Loop through each slot again to check for clashes too near to current slot
                            While (m < numSlots)
                                If (Abs(slots(l).day - slots(m).day) <= minDutyGap And personnel(i).name = slots(m).personnel) Then
                                    hasClash = True
                                End If
                                m = m + 1
                            Wend
                            
                            ' If the personnel has committment on that particular day,
                            ' then hasClash = True
                            If (personnel(i).getCommitment(slots(l).day)) Then hasClash = True
                            
                            ' If the personnel is able to take arms And
                            ' the slot is an armed slot,
                            ' then hasClash = False
                            If (Not personnel(i).armed And slots(l).armed) Then
                                hasClash = True
                            End If
                            
                            ' If Not hasClash(False) = True; then
                            ' declare that a slot has been found and assign
                            ' the personnel to that particular slot.
                            ' Next, shuffleSlotsByDifficulty.
                            If (Not hasClash) Then
                                foundSlot = True
                                slots(l).personnel = personnel(i).name
                                shuffleSlotsByDifficulty
                            End If
                        End If
                        l = l + 1
                    Wend
                    
                    ' TODO : I don't understand this part.
                    ' If slots are not found, Then hasMissingSlot = True
                    ' log to Guard Duty planner sheet
                    ' increment numErrors by 1
                    If (Not foundSlot) Then
                        hasMissingSlot = True
                        ' log.log ("Unable to assign " & personnel(i).name & " to a " & j & " point slot")
                        errorLog(numErrors) = "Unable to assign " & personnel(i).name & " to a " & j & " point slot"
                        numErrors = numErrors + 1
                    End If
                    
                    k = k + 1
                Wend
                j = j + 1
            Wend
            i = i + 1
        Wend
        
        ' Assign standby
        ' This is the exact same code as before. The comments written previously apply here too.
        '
        i = 0
        While (i < numPersonnel) ' Loop through each of 41 personnel
            j = 1
            'Initially, numDutiesToAssign is 0.
            'Subsequently, numDutiesToAssign is
            numDutiesToAssign = personnel(i).numberOfStandbys
            
           'While (1 <= 2)
            While (j <= 2) ' Loop through each duty points
                ' k is an iterative variable.
                k = 0
                While (k < numDutiesToAssign(j)) ' Loop through each 1 point / 2 points duty slot.
                    l = 0
                    foundSlot = False
                    While (l < numSlots) ' Loop through each slot
                        
                        ' Check if the slot is vacant or if there is a solution.
                        If (Not foundSlot And slots(l).standby = "" And slots(l).points = j) Then
                            m = 0
                            hasClash = False
                            While (m < numSlots) ' Loop through each slot again to check for clashes with current slot
                                If (Abs(slots(l).day - slots(m).day) <= minStbGap And personnel(i).name = slots(m).personnel) Then
                                    hasClash = True
                                End If
                                m = m + 1
                            Wend
                            
                            m = 0
                            
                            ' Loop through each slot again to check for clashes w/ standby too near to current slot
                            While (m < numSlots)
                                If (Abs(slots(l).day - slots(m).day) <= minStbGap And personnel(i).name = slots(m).standby) Then
                                    hasClash = True
                                End If
                                m = m + 1
                            Wend
                            
                            If (personnel(i).getCommitment(slots(l).day)) Then hasClash = True
                            
                            If (Not personnel(i).armed And slots(l).armed) Then
                                hasClash = True
                            End If
                            
                            If (Not hasClash) Then
                                foundSlot = True
                                slots(l).standby = personnel(i).name
                                shuffleSlotsByDifficulty
                            End If
                        End If
                        l = l + 1
                    Wend
                    
                    If (Not foundSlot) Then
                        hasMissingSlot = True
                        ' log.log ("Unable to assign " & personnel(i).name & " to a " & j & " point standby slot")
                        errorLog(numErrors) = "Unable to assign " & personnel(i).name & " to a " & j & " point standby slot"
                        numErrors = numErrors + 1
                    End If
                    
                    k = k + 1
                Wend
                j = j + 1
            Wend
            i = i + 1
        Wend
    Wend
    
    ' log.log ("Retried " & numRetries & " times")
    If (hasMissingSlot) Then
        i = 0
        While (i < numErrors)
            log.log (errorLog(i))
            i = i + 1
        Wend
    End If
    
    i = 0
    While (i < numSlots) ' While (i < 52)
        slots(i).writeToDutyList
        slots(i).HighlightEmpty
        'log.log (i & "- " & slots(i).toString)
        i = i + 1
    Wend
    
    ' This subroutine is useless.
    i = 0
    While (i < numPersonnel)
        'log.log (i & "- " & personnel(i).toString)
        i = i + 1
    Wend
    
    ResetHighlightCommitments
    
    ' POINT SYSTEM
    ' Existing points (lower ppm, higher points)
    ' Duty commitments (more days away, higher points)
    ' Armed (non-armed +++++++ points make sure put first)
    
    
    
    ' Assign points to everyone based on their PPM (points per month)
End Sub

' Add duty records from the filled duty slots into the duty records sheet
Sub AddDutyRecords()
    
    initVars
    
    Dim numRows As Integer
    Dim currentRow As Integer
    Dim currentCol As Integer
    Dim checkCol As Integer
    
    Dim hasClash As Boolean
    
    Dim i As Integer, j As Integer
    
    Dim dayHasDuty As Boolean
    
    numRows = countRows
    
    ' Loop through each row
    i = 0
    
    While (i < numRows)
        currentRow = dutySlotsStartRow + i
        ' If cell is not black
        dayHasDuty = False
        j = 0
        While (j < numDutyCols)
            If (Worksheets(dutySlotsSheet).Cells(currentRow, firstActualCol + 2 * j).Interior.Color <> RGB(0, 0, 0)) Then dayHasDuty = True
            j = j + 1
        Wend
        If (dayHasDuty) Then
            j = 0
            While (j < numDutyCols)
                If (Worksheets(dutySlotsSheet).Cells(currentRow, firstActualCol + 2 * j).Interior.Color <> RGB(0, 0, 0)) Then
                    DutyRecords.AddDutyRecord Worksheets(dutySlotsSheet).Cells(currentRow, firstActualCol + 2 * j).Value, Worksheets(dutySlotsSheet).Range(monthCell).Value, dutyType, Worksheets(dutySlotsSheet).Cells(currentRow, pointsCol).Value
                End If
                j = j + 1
            Wend
            
        End If
        
        i = i + 1
    Wend

End Sub

' Check for errors in the current duty slots
Sub CheckForErrors()
    
    initVars
    
    Dim numRows As Integer
    Dim currentRow As Integer
    Dim currentCol As Integer
    Dim checkCol As Integer
    
    Dim hasClash As Boolean
    
    Dim dayHasDuty As Boolean
    
    Dim i As Integer, j As Integer
    
    numRows = countRows
    
    ' Loop through each row
    i = 0
    While (i < numRows)
        currentRow = dutySlotsStartRow + i
        
        ' Keep track of clashes
        hasClash = False
        
        dayHasDuty = False
        j = 0
        While (j < numDutyCols)
            If (Worksheets(dutySlotsSheet).Cells(currentRow, firstActualCol + 2 * j).Interior.Color <> RGB(0, 0, 0)) Then dayHasDuty = True
            j = j + 1
        Wend
        
        ' If cell is not black
        If (dayHasDuty) Then
            
            ' Define checking range
            Dim checkRowFirst As Integer
            Dim checkRow As Integer
            checkRowFirst = currentRow - 2
            If (checkRowFirst < dutySlotsStartRow) Then
                checkRowFirst = dutySlotsStartRow
            End If
            Dim checkRowLast As Integer
            checkRowLast = currentRow + 2
            
            ' Check all 4 col
            currentCol = firstActualCol
            While (currentCol < pointsCol)
            
                ' Check that the cell we're checking is not blank
                If (Worksheets(dutySlotsSheet).Cells(currentRow, currentCol).Value <> "") Then
                    ' Loop through each row and col
                    checkCol = firstCheckCol
                    While (checkCol < pointsCol)
                        ' Loop through +- 2 from current row
                        checkRow = checkRowFirst
                        While (checkRow < checkRowLast + 1)
                            'MsgBox ("checking row " & checkRow & " col " & checkCol)
                                
                            ' Don't check against itself
                            If Not (currentRow = checkRow And currentCol = checkCol) Then
                                ' Check for clash
                                If (Worksheets(dutySlotsSheet).Cells(currentRow, currentCol).Value = Worksheets(dutySlotsSheet).Cells(checkRow, checkCol).Value) Then
                                    hasClash = True
                                    'MsgBox ("CLASHING " & currentRow & ", " & currentCol & " with " & checkRow & ", " & checkCol)
                                    Worksheets(dutySlotsSheet).Cells(currentRow, currentCol).Interior.Color = RGB(255, 255, 0)
                                End If
                            End If
                            
                            checkRow = checkRow + 1
                        Wend
                        checkCol = checkCol + 1
                    Wend
                End If
                currentCol = currentCol + 1
            Wend
        End If
        
        If (hasClash) Then
            'MsgBox ("Ouch! There's a clash at row " & currentRow)
        End If
        i = i + 1
    Wend

End Sub

Sub FindUncommitted()

    initVars
    
    Dim selectedDate As Integer
    Dim output As String
    
    Dim i As Integer
    
    selectedDate = Worksheets(dutySlotsSheet).Cells(ActiveCell.Row, dateCol).Value
    
    ' Load everyone
    
    Dim n As Integer
    n = PointsTable.countRows
    i = 0
    
    output = "Personnel without commitments:" & vbCrLf
    
    Dim dp As DutyPersonnel
    While (i < n)
        ' Don't count with exemptions
        If (Not DutyExemptions.PersonnelHasExemption(PointsTable.getName(i + 2)) And PointsTable.getDutyType(i + 2) = dutyType) Then
            Set personnel(numPersonnel) = New DutyPersonnel
            personnel(numPersonnel).initialize (PointsTable.getName(i + 2))
            
            If (Not personnel(numPersonnel).getCommitment(selectedDate)) Then
                output = output & personnel(numPersonnel).name & vbCrLf
            End If
            
            numPersonnel = numPersonnel + 1
        End If
        i = i + 1
    Wend
    
    MsgBox (output)
    
End Sub

Sub HighlightCommitments()
    initVars
    Dim i As Integer, j As Integer
    
    Dim selectedPers As String
    Dim pers As DutyPersonnel
    Dim currentDate As Integer
    
    ResetHighlightCommitments
    
    selectedPers = Worksheets(dutySlotsSheet).Cells(ActiveCell.Row, ActiveCell.Column).Value
    Set pers = New DutyPersonnel
    pers.initialize (selectedPers)
    
    i = dutySlotsStartRow
    While (Worksheets(dutySlotsSheet).Cells(i, dateCol).Value <> "")
        currentDate = Worksheets(dutySlotsSheet).Cells(i, dateCol).Value
        If (pers.getCommitment(currentDate)) Then
            j = firstActualCol
            While (j < pointsCol)
                If (Worksheets(dutySlotsSheet).Cells(i, j).Interior.ColorIndex <> 1) Then
                    Worksheets(dutySlotsSheet).Cells(i, j).Interior.ColorIndex = 27
                End If
                j = j + 1
            Wend
        End If
        i = i + 1
    Wend
End Sub

Sub HighlightDuties()

    initVars
    Dim i As Integer, j As Integer
    
    Dim selectedPers As String
    
    ResetHighlightCommitments
    
    selectedPers = Worksheets(dutySlotsSheet).Cells(ActiveCell.Row, ActiveCell.Column).Value
    
    i = dutySlotsStartRow
    While (Worksheets(dutySlotsSheet).Cells(i, dateCol).Value <> "")
        j = firstActualCol
        While (j < pointsCol)
            If (Worksheets(dutySlotsSheet).Cells(i, j).Interior.ColorIndex <> 1) Then
                If (Worksheets(dutySlotsSheet).Cells(i, j).Value = selectedPers) Then
                    Worksheets(dutySlotsSheet).Cells(i, j).Interior.ColorIndex = 27
                End If
            End If
            j = j + 1
        Wend
        i = i + 1
    Wend
End Sub

Sub ResetHighlightCommitments()
    initVars
    Dim i As Integer, j As Integer
    
    Dim currentDay As String
    
    Dim colo As Integer
    
    i = dutySlotsStartRow
    While (Worksheets(dutySlotsSheet).Cells(i, dateCol).Value <> "")
        currentDay = Worksheets(dutySlotsSheet).Cells(i, dayCol).Value
        j = firstActualCol
        While (j < pointsCol)
            If (Worksheets(dutySlotsSheet).Cells(i, j).Interior.ColorIndex <> 1) Then
                If (currentDay = "SAT" Or currentDay = "SUN") Then
                    Worksheets(dutySlotsSheet).Cells(i, j).Interior.ColorIndex = 15
                Else
                    Worksheets(dutySlotsSheet).Cells(i, j).Interior.ColorIndex = 0
                End If
                
                If (Worksheets(dutySlotsSheet).Cells(i, j).Value = "") Then
                
                    If (currentDay = "SAT" Or currentDay = "SUN") Then
                        colo = 28
                    Else
                        colo = 33
                    End If
                
                    Worksheets(dutySlotsSheet).Cells(i, j).Interior.ColorIndex = colo
                End If
            End If
            j = j + 1
        Wend
        i = i + 1
    Wend
End Sub

Private Sub shuffleSlotsByDifficulty()
    
    ' A brief explanation:
    ' I am guessing this subroutine randomly shuffles the slots.
    
    Dim i As Integer, j As Integer, k As Integer
    Dim tempslot As DutySlot
    Dim numDutiesWithDiff As Integer
    Dim currDiff As Double
    Dim toSwap1 As Integer
    Dim toSwap2 As Integer
    
    ' Call sortSlotsByDifficulty to arrange the slots based on the difficulty value.
    sortSlotsByDifficulty
    
    i = 2
    j = 0
    While (j < numSlots) ' While (j < 52)
        
        numDutiesWithDiff = 0
        
        k = j ' k = 0
        currDiff = slots(k).difficulty
        
        While (k < numSlots) ' While (0 < 52)
            If (slots(k).difficulty = currDiff) Then
                numDutiesWithDiff = numDutiesWithDiff + 1
            Else
                k = numSlots ' k = 52
            End If
            k = k + 1
        Wend
        
        k = 0
        While (k < 100)
            toSwap1 = j + CInt(Int(numDutiesWithDiff * Rnd()))
            toSwap2 = j + CInt(Int(numDutiesWithDiff * Rnd()))
            
            Set tempslot = slots(toSwap1)
            Set slots(toSwap1) = slots(toSwap2)
            Set slots(toSwap2) = tempslot
            k = k + 1
        Wend
        
        j = j + numDutiesWithDiff
    Wend
End Sub

Private Sub sortPersonnelByPoints()
    ' This subroutine sorts the personnel by points in a linear fashion.
    Dim i As Integer, j As Integer
    Dim highestIndex As Integer
    Dim highestPoints As Double
    
    ' Create tempPersonnel object with DutyPersonnel class module
    Dim tempPersonnel As DutyPersonnel
    
    ' Sort personnel by points
    i = 0
    While (i < numPersonnel - 1) ' While (i < 41 - 1)
        j = i ' j = 0
        
        highestIndex = 0
        highestPoints = -1
        
        While (j < numPersonnel) ' While (j < 41)
            If (personnel(j).dutyPoints > highestPoints) Then
                highestPoints = personnel(j).dutyPoints
                highestIndex = j
            End If
            j = j + 1
        Wend
        
        Set tempPersonnel = personnel(i)
        Set personnel(i) = personnel(highestIndex)
        Set personnel(highestIndex) = tempPersonnel
        i = i + 1
    Wend
End Sub

Private Sub sortPersonnelByPointsReverse()
    Dim i As Integer, j As Integer
    Dim highestIndex As Integer
    Dim highestPoints As Double
    Dim tempPersonnel As DutyPersonnel
    
    ' Sort personnel by points
    i = 0
    While (i < numPersonnel - 1)
        j = i
        highestIndex = 0
        highestPoints = 999
        While (j < numPersonnel)
            If (personnel(j).dutyPoints < highestPoints) Then
                highestPoints = personnel(j).dutyPoints
                highestIndex = j
            End If
            j = j + 1
        Wend
        Set tempPersonnel = personnel(i)
        Set personnel(i) = personnel(highestIndex)
        Set personnel(highestIndex) = tempPersonnel
        i = i + 1
    Wend
End Sub

Private Sub sortPersonnelByDifficulty()
    
    ' Iterative Variables
    Dim i As Integer, j As Integer
    
    Dim highestIndex As Integer
    Dim highestPoints As Integer
    
    Dim tempPersonnel As DutyPersonnel
    
    ' Sort personnel by difficulty aka
    i = 0
    While (i < numPersonnel - 1) ' While (i < 41 - 1)
        j = i ' j = i = 0
        
        highestIndex = 0
        highestPoints = -1
        
        While (j < numPersonnel) ' While (j < 41)
            ' personnel(?).difficulty is the total value of
            ' white cells present (the personnel has no committments for that following day)
            ' in the personnel's row in committments sheet.
            ' The personnel who has many committments will have a higher number
            ' than a personnel who has little committments.
            If (personnel(j).difficulty > highestPoints) Then
                highestPoints = personnel(j).difficulty
                highestIndex = j
            End If
            j = j + 1
        Wend
        Set tempPersonnel = personnel(i)
        Set personnel(i) = personnel(highestIndex)
        Set personnel(highestIndex) = tempPersonnel
        i = i + 1
    Wend
End Sub

Private Sub sortSlotsByDifficulty()
    
    ' ============ A brief explanation on what this subroutine does ============
    ' 1 slot is defined as a column for a particular date, for example
    ' 1st of Nov, 2nd of Nov, 3rd of Nov and so on.
    ' TODO
    ' In this column D for example, there are empty, P, SB slots from all 41 guards.
    ' Therefore, this subroutine sorts all 52 slots based on the availability of the guards
    ' in a decreasing order.
    ' If the availability of the guards is very low for that column, that slot will have a higher difficulty value.
    
    
    Dim highestDiff As Integer
    Dim tempslot As DutySlot
    Dim highestIndex As Integer
    
    Dim i As Integer, j As Integer
    
    i = 0
    While (i < numSlots - 1) ' While (0 .. 50 < 52 - 1)
        j = i ' j = i = 0
        
        highestIndex = 0
        highestDiff = -1
        
        While (j < numSlots) ' While (0 ... 51 < 52)
            
            If (slots(j).difficulty > highestDiff) Then
                highestDiff = slots(j).difficulty
                highestIndex = j
            End If
            j = j + 1
        Wend
        
        Set tempslot = slots(i)
        Set slots(i) = slots(highestIndex)
        Set slots(highestIndex) = tempslot
        i = i + 1
    Wend
End Sub

Private Sub sortSlotsByPoints()
    Dim highestDiff As Integer
    Dim tempslot As DutySlot
    Dim highestIndex As Integer
    
    ' Iterative variables
    Dim i As Integer, j As Integer
    
    i = 0
    While (i < numSlots - 1) ' While (i < 52 - 1)
        j = i ' j = 0
        highestIndex = 0
        highestDiff = -1
        
        ' This portion of code select the slot which has the highest number of points.
        While (j < numSlots) ' While (j < 52)
            If (slots(j).points > highestDiff) Then ' If (slots(0).points > highestDiff) Then
                highestDiff = slots(j).points
                highestIndex = j
            End If
            j = j + 1
        Wend
        
        ' this is the original slots(0)
        Set tempslot = slots(i)
        ' Replace original slots(0) with new slots(highestIndex)
        Set slots(i) = slots(highestIndex)
        ' Place the tempSlot in the slots(highestIndex)
        Set slots(highestIndex) = tempslot
        i = i + 1
    Wend
End Sub

Private Sub initVars()
    dutySlotsSheet = "Duty Slots"
    dutySlotsStartRow = 3
    
    dateCol = 1
    dayCol = 2
    firstActualCol = 3
    firstStbCol = 4
    
    numDutyCols = 0
    Dim i As Integer
    i = 3
    
    While (Worksheets(dutySlotsSheet).Cells(2, i).Value <> "POINTS")
        numDutyCols = numDutyCols + 1 ' numDutyCols = 2
        i = i + 2
    Wend
    
    pointsCol = i ' pointsCol = 7
    
    firstCheckCol = 3
    lastCheckCol = 6
    
    dutyTypeCell = "C1" ' Checks if its for Guard or Guard 2 IC
    monthCell = "D1"
    dutyHeaderRow = 2
    
    dutyType = Worksheets(dutySlotsSheet).Range(dutyTypeCell).Value
    planningmonth = Worksheets(dutySlotsSheet).Range(monthCell).Value
    
    plannerSheet = "Guard Duty Planner"
    dutyGapCell = "M4"
    standbyGapCell = "M5"
    minDutyGap = Worksheets(plannerSheet).Range(dutyGapCell).Value
    minStbGap = Worksheets(plannerSheet).Range(standbyGapCell).Value
    
End Sub

Function countRows()
    Dim i As Integer
    
    i = dutySlotsStartRow
    While (Not IsEmpty(Worksheets(dutySlotsSheet).Cells(i, 1).Value))
        i = i + 1
    Wend
    countRows = i - dutySlotsStartRow
    ' The return value is 52
End Function

Function getPlanningMonth()
    initVars
    getPlanningMonth = planningmonth
End Function

Function getDutyType()
    initVars
    getDutyType = Worksheets(dutySlotsSheet).Range(dutyTypeCell).Value
End Function

Function getPointsCol()
    getPointsCol = pointsCol
End Function

Function getColHeader(col As Integer)
    getColHeader = Worksheets(dutySlotsSheet).Cells(dutyHeaderRow, col).Value
End Function



