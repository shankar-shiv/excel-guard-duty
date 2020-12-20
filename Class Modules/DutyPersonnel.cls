Rem Attribute VBA_ModuleType=VBAClassModule

Option Explicit

Dim detailsSn As Integer
Dim commitmentsSn As Integer
Dim pRank As String
Dim pName As String
Dim pContactNo As String
Dim pDifficulty As Integer
Dim pDutyPoints As Double
Dim pArms As Boolean

Dim detailsSheet As String
Dim commitmentsSheet As String

Dim rankCol As Integer
Dim nameCol As Integer
Dim contactNoCol As Integer
Dim ppmCol As Integer
Dim armsCol As Integer

Dim commitmentsNameCol As Integer
Dim commitmentsStartCol As Integer
Dim commitmentsDateRow As Integer

Dim detailsStartRow As Integer
Dim commitmentsStartRow As Integer
Dim numDuties(2) As Integer
Dim numStandbys(2) As Integer
Dim dutyDays(32) As Integer
Dim numDutyDays As Integer

Private Sub Class_Initialize()
    detailsSheet = "Points Table"
    commitmentsSheet = "Duty Commitments"
    rankCol = 2
    nameCol = 3
    contactNoCol = 4
    ppmCol = 9
    armsCol = 8
    detailsStartRow = 2
    commitmentsStartRow = 3
    commitmentsNameCol = 2 ' Original : 4
    commitmentsStartCol = 6 ' Might be correct, i think it's 4.
    commitmentsDateRow = 2
    
    pDifficulty = 0
    numDuties(0) = 0
    numDuties(1) = 0
    numDuties(2) = 0
    numStandbys(0) = 0
    numStandbys(1) = 0
    numStandbys(2) = 0
    numDutyDays = 0
End Sub

Sub initialize(n As String) ' n is the name of the guard
    Dim rowNum As Integer
    rowNum = detailsStartRow ' rowNum = 2
                                 
                                 ' .Cells(2, 3)
    While (Worksheets(detailsSheet).Cells(rowNum, nameCol).Value <> "")
        If (Worksheets(detailsSheet).Cells(rowNum, nameCol).Value = n) Then
            detailsSn = rowNum
            pName = n
            pRank = Worksheets(detailsSheet).Cells(rowNum, rankCol).Value
            pContactNo = Worksheets(detailsSheet).Cells(rowNum, contactNoCol).Value
            
                ' Round() returns a number rounded to a specified number of decimal places
            pDutyPoints = Round(Worksheets(detailsSheet).Cells(rowNum, ppmCol).Value, 3)
            If (Worksheets(detailsSheet).Cells(rowNum, armsCol).Value = "Y") Then
                pArms = True
            Else
                pArms = False
            End If
        End If
        rowNum = rowNum + 1
    Wend
    
    ' Commitments
    rowNum = commitmentsStartRow ' rowNum = commitmentsStartRow = 3
                                     
                                     ' .Cells(3, 2)
    While (Worksheets(commitmentsSheet).Cells(rowNum, commitmentsNameCol).Value <> "")
        If (Worksheets(commitmentsSheet).Cells(rowNum, commitmentsNameCol).Value = n) Then
            commitmentsSn = rowNum
        End If
        rowNum = rowNum + 1
    Wend
    
    ' Calculate Difficulty for each guards.
    Dim i As Integer
    ' Why is commitmentsStartCol = 6, shouldn't it be 4
    i = commitmentsStartCol ' i = commitmentsStartCol = 6
                   
    ' This While...Wend loop calculates a difficulty score for each guard.
    ' What's a difficulty score ?
    
    ' Well, let's understand some variables. commitmentsSn is the row number of the guards.
    ' It will range from 3 to 43.
    ' If the cell (row=3 to 43, column=6 to 30 days of the month) is Not empty, which means the
    ' personnel has some kind of committment on that particular day, the pDifficulty value would increase by 1.
    
                                     ' .Cells(        2         , 6)
    While (Worksheets(commitmentsSheet).Cells(commitmentsDateRow, i).Value <> "")
        If (Worksheets(commitmentsSheet).Cells(commitmentsSn, i).Value <> "") Then
            pDifficulty = pDifficulty + 1
        End If
        i = i + 1
    Wend
    
    ' If the guard cannot take up arms, increase pDifficulty to the max.
    ' This is to ensure that the guard isn't assigned for any duty.
    If (Not pArms) Then pDifficulty = pDifficulty + 50
    
End Sub

Function sortByPoints(ByRef personnel() As DutyPersonnel) As DutyPersonnel()
    Dim numPersonnel As Integer
    numPersonnel = PointsTable.countRows
    
    Dim i As Integer
    i = 0
    
End Function

Sub addDuty(points As Integer)
    numDuties(points) = numDuties(points) + 1
End Sub

Sub removeDuty(points As Integer)
    numDuties(points) = numDuties(points) - 1
    If (numDuties(points) < 0) Then numDuties(points) = 0
End Sub

Sub addDutyDay(day As Integer)
    If (Not getDutyDay(day)) Then
        dutyDays(numDutyDays) = day
        numDutyDays = numDutyDays + 1
    End If
End Sub

Function getDutyDay(day As Integer) As Boolean
    Dim i As Integer
    Dim have As Boolean
    i = 0
    While (i < numDutyDays)
        If (dutyDays(i) = day) Then
            have = True
        End If
        i = i + 1
    Wend
    If (have) Then
        getDutyDay = True
    Else
        getDutyDay = False
    End If
End Function

Sub addStandby(points As Integer)
    numStandbys(points) = numStandbys(points) + 1
End Sub

Sub removeStandby(points As Integer)
    numStandbys(points) = numStandbys(points) - 1
    If (numStandbys(points) < 0) Then numStandbys(points) = 0
End Sub

Function getCommitment(day As Integer) As Boolean
    ' If not blank or V(volunteer)
    If (Worksheets(commitmentsSheet).Cells(commitmentsSn, day + commitmentsStartCol - 1).Value <> "" And Worksheets(commitmentsSheet).Cells(commitmentsSn, day + commitmentsStartCol - 1).Value <> "V") Then
        getCommitment = True
    Else
        getCommitment = False
    End If
End Function

Function getVolunteer(day As Integer) As Boolean
    ' If not blank or V(volunteer)
    If (Worksheets(commitmentsSheet).Cells(commitmentsSn, day + commitmentsStartCol - 1).Value = "V") Then
        getVolunteer = True
    Else
        getVolunteer = False
    End If
End Function

Function toString()
    Dim output As String
    output = detailsSn & ", " & commitmentsSn & ": " & pRank & ", " & pName & ", " & pContactNo & ", DIF: " & pDifficulty & ", PTS: " & pDutyPoints & ", Duties: "
    
    Dim i As Integer, j As Integer
    i = 1
    While (i < 3)
        j = 0
        While (j < numDuties(i))
            output = output & i & ", "
            j = j + 1
        Wend
        i = i + 1
    Wend
    
    output = output & " STB:"
    
    i = 1
    While (i < 3)
        j = 0
        While (j < numStandbys(i))
            output = output & i & ", "
            j = j + 1
        Wend
        i = i + 1
    Wend
    
    toString = output
End Function

Public Property Get serialNo() As String
    serialNo = detailsSn
End Property

Public Property Get rank() As String
    rank = pRank
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Get contactNo() As String
    contactNo = pContactNo
End Property

Public Property Get difficulty() As String
    difficulty = pDifficulty
End Property

Public Property Get dutyPoints() As String
    dutyPoints = pDutyPoints
End Property

Public Property Get armed() As Boolean
    armed = pArms
End Property

Public Property Get numberOfDuties() As Integer()
    numberOfDuties = numDuties
End Property

Public Property Get numberOfStandbys() As Integer()
    numberOfStandbys = numStandbys
End Property

Public Property Get totalPoints() As Integer
    totalPoints = numDuties(1) + numDuties(2) * 2
End Property

Public Property Get numberOfDutiesWithPoints(p As Integer) As Integer
    numberOfDutiesWithPoints = numDuties(p)
End Property

Public Property Get numberOfStandbysWithPoints(p As Integer) As Integer
    numberOfStandbysWithPoints = numStandbys(p)
End Property



