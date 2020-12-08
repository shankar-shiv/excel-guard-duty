' Automate planning of duties
Sub PlanDuties()
    
    'Call the sub routine initVars() to initialise variables
    initVars
    
    Dim numRows As Integer
    Dim currentRow As Integer
    
    Dim totalPoints As Integer
    totalPoints = 0
    Dim numDuties As Integer
    numDuties = 0
    Dim overUnder As Double
    
    ' Track unfulfilled duties to swap
    
    
    Dim cont As Boolean
    
    Dim ds As DutySlot
    
    Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer
    ' Dim i, j, k, l, m As Integer
    
    Dim slotPointsBackup(2) As Integer
    
    numPersonnel = 0
    numSlots = 0
    
    slotPoints(0) = 0
    slotPoints(1) = 0
    slotPoints(2) = 0
    
    numRows = countRows
    
    Set log = New Logger ' Initialise log object
    log.clearLog
    
    'log.log ("Number of duties per day: " & numDutyCols)
    
    Dim dayHasDuty As Boolean
    
    ' Count total points, number of slots, load all slots
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
                If (Worksheets(dutySlotsSheet).Cells(currentRow, firstActualCol + 2 * j).Interior.Color <> RGB(0, 0, 0)) Then ' Unarmed
                    slotPoint = Worksheets(dutySlotsSheet).Cells(currentRow, pointsCol).Value
                    totalPoints = totalPoints + slotPoint
                    numDuties = numDuties + 1
                    slotPoints(slotPoint) = slotPoints(slotPoint) + 1
                    
                    Set slots(numSlots) = New DutySlot
                    Call slots(numSlots).initialize(currentRow, firstActualCol + 2 * j)
                    numSlots = numSlots + 1
                End If
                j = j + 1
            Wend
            
        End If
        
        i = i + 1
    Wend
    
    i = 0
    While (i < 3)
        slotPointsBackup(i) = slotPoints(i)
        i = i + 1
    Wend
    
    overUnder = totalPoints / numDuties
    
    log.log ("Total Points: " & totalPoints)
    log.log ("Total Duties: " & numDuties)
    'log.log ("Avg Points Per Slot: " & overUnder)
    
    ' Load everyone
    Dim n As Integer
    n = PointsTable.countRows
    i = 0
    
    Dim dp As DutyPersonnel
    While (i < n)
        ' Don't count with exemptions
        If (Not DutyExemptions.PersonnelHasExemption(PointsTable.getName(i + 2)) And PointsTable.getDutyType(i + 2) = dutyType) Then
            Set personnel(numPersonnel) = New DutyPersonnel
            personnel(numPersonnel).initialize (PointsTable.getName(i + 2))
            
            numPersonnel = numPersonnel + 1
        End If
        i = i + 1
    Wend
    
    log.log ("Number of Duty Personnel: " & numPersonnel)
    log.log ("Points Per Person: " & (totalPoints / numPersonnel) & vbCrLf)
    
    ' Sort personnel by points; call the sub routine sortPersonnelByPoints()
    sortPersonnelByPoints
    
    ' Remove extras from assigning list
    slotPoints(2) = slotPoints(2) - DutyExtras.CountTotalMonthExtras
    If (slotPoints(2) < 0) Then slotPoints(2) = 0
    
    ' Setup slots
    Dim currentDay As Integer
    Dim numVolunteers As Integer
    numVolunteers = 0
    
    i = 0
    While (i < numSlots)
        currentDay = slots(i).day
        j = 0
        While (j < numPersonnel)
            If (personnel(j).getCommitment(currentDay)) Then ' Set difficulty
                slots(i).difficulty = slots(i).difficulty + 1
            End If
            
            If (personnel(j).getVolunteer(currentDay) And Not personnel(j).getDutyDay(currentDay)) Then ' Set volunteer
                                
                slots(i).setVolunteer (personnel(j).name)
                personnel(j).addDutyDay (currentDay)
            End If
            j = j + 1
        Wend
        
        If (slots(i).personnel <> "") Then ' Pre-allocated / Volunteer personnel
            j = 0
            While (j < numPersonnel)
                If (personnel(j).name = slots(i).personnel) Then
                    slots(i).setVolunteer (personnel(j).name)
                    
                    slotPoints(slots(i).points) = slotPoints(slots(i).points) - 1
                    numVolunteers = numVolunteers + 1
                    
                    If (personnel(j).getVolunteer(slots(i).day)) Then
                        log.log (personnel(j).name & " volunteered on " & currentDay)
                    Else
                        log.log (personnel(j).name & " was pre-assigned on " & currentDay)
                    End If
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
    Dim cpp As Integer ' Current Planning Points (points of the slot)
    Dim currIndex As Integer
    Dim breakPoint As Integer
    breakPoint = -1
    cpp = 2
    currIndex = numPersonnel - 1
    i = 0
    
    While (i < numDuties - DutyExtras.CountTotalMonthExtras - numVolunteers)
        If (cpp = 2) Then
            personnel(currIndex).addDuty (cpp)
            slotPoints(cpp) = slotPoints(cpp) - 1
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
        
        If (slotPoints(cpp) = 0) Then
            cpp = cpp - 1
            If (cpp = 1 And currIndex > 0) Then breakPoint = currIndex - 1
        End If
        
        currIndex = currIndex - 1
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
    currIndex = numPersonnel - 1
    i = 0
    
    While (cont)
    ' Both personnel should have the same number of total points
        If (personnel(currIndex).totalPoints = personnel(i).totalPoints And personnel(i).dutyPoints > 0 And personnel(currIndex).dutyPoints > 0) Then
            ' personnel(i).dutyPoints > 0.1 => returns a negative -1 number
            If ((personnel(i).dutyPoints - personnel(currIndex).dutyPoints) / personnel(i).dutyPoints > 0.1) Then
                If (personnel(i).numberOfDutiesWithPoints(1) > 0) Then
                    personnel(i).removeDuty (1)
                    personnel(currIndex).addDuty (1)
                End If
            End If
        Else
            cont = False
        End If
        currIndex = currIndex - 1
        i = i + 1
        
    Wend
    
    ' Assign extras
    i = 0
    Dim numExtras As Integer
    
    While (i < numPersonnel)
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
    
    sortSlotsByPoints
    
    ' Sort slots and personnel by difficulty
    
    sortPersonnelByDifficulty
    
    log.log ("")
    
    'Assign duties
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
    
    While (hasMissingSlot And numRetries < 1)
        hasMissingSlot = False
        numRetries = numRetries + 1
        numErrors = 0
        
        i = 0
        While (i < numSlots)
            If (Not slots(i).locked) Then
                slots(i).personnel = ""
                slots(i).standby = ""
            End If
            i = i + 1
        Wend
        
        i = 0
        While (i < numPersonnel) ' Loop through each personnel
            j = 1
            
            numDutiesToAssign = personnel(i).numberOfDuties
            ' how many duties a person has ex : John = 2 duties
            
            
            Randomize
            Dim MyValue As Integer
            MyValue = Int((5 - 2 + 1) * Rnd + 2)
            
            While (j <= 2) ' Loop through each duty points
                k = 0
                While (k < numDutiesToAssign(j)) ' Loop through each duty
                    l = 0
                    foundSlot = False
                    While (l < MyValue) ' Loop through each slot : numSlots
                        If (Not foundSlot And slots(l).personnel = "" And slots(l).points = j) Then ' Check is slot is open or found solution
                            m = 0
                            hasClash = False
                            While (m < numSlots) ' Loop through each slot again to check for clashes too near to current slot
                                If (Abs(slots(l).day - slots(m).day) <= minDutyGap And personnel(i).name = slots(m).personnel) Then
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
                                slots(l).personnel = personnel(i).name
                                shuffleSlotsByDifficulty
                            End If
                        End If
                        l = l + 1
                    Wend
                    
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
        i = 0
        While (i < numPersonnel) ' Loop through each personnel
            j = 1
            numDutiesToAssign = personnel(i).numberOfStandbys
            While (j <= 2) ' Loop through each duty points
                k = 0
                While (k < numDutiesToAssign(j)) ' Loop through each duty
                    l = 0
                    foundSlot = False
                    While (l < numSlots) ' Loop through each slot
                        If (Not foundSlot And slots(l).standby = "" And slots(l).points = j) Then ' Check is slot is open or found solution
                            m = 0
                            hasClash = False
                            While (m < numSlots) ' Loop through each slot again to check for clashes with current slot
                                If (Abs(slots(l).day - slots(m).day) <= minStbGap And personnel(i).name = slots(m).personnel) Then
                                    hasClash = True
                                End If
                                m = m + 1
                            Wend
                            
                            m = 0
                            While (m < numSlots) ' Loop through each slot again to check for clashes w/ standby too near to current slot
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
    While (i < numSlots)
        slots(i).writeToDutyList
        slots(i).HighlightEmpty
        'log.log (i & "- " & slots(i).toString)
        i = i + 1
    Wend
    
    i = 0
    While (i < numPersonnel)
        'log.log (i & "- " & personnel(i).toString)
        i = i + 1
    Wend
    
    ResetHighlightCommitments
    
    ' POINT SYSTEM
    ' Existing points : (lower points per month (ppm) , higher points)
    ' Duty commitments (more days away, higher points)
    ' Armed (non-armed +++++++ points make sure put first)
    
    
    
    ' Assign points to everyone based on their PPM
End Sub

