Rem Attribute VBA_ModuleType=VBAClassModule

Option Explicit

Dim sPoints As Integer
Dim sRow As Integer
Dim sCol As Integer
Dim sStbCol As Integer
Dim sDay As Integer
Dim sPersonnel As String
Dim sStandby As String
Dim sDifficulty As Integer
Dim sLocked As Boolean
Dim sArmed As Boolean

Dim slotsSheet As String
Dim dayCol As Integer
Dim pointsCol As Integer

Private Sub Class_Initialize()
    slotsSheet = "Duty Slots"
    dayCol = 1 ' dayCol should be 2
    pointsCol = DutySlots.getPointsCol ' DutySlot is a Class Module, DutySlots is a Module
    sPersonnel = ""
    sStandby = ""
    sDifficulty = 0
    sLocked = False
End Sub

Sub initialize(r As Integer, c As Integer)
    sRow = r
    sCol = c
    sStbCol = c + 1
    ' This is actually the date of the slot NOT the day, very confusing
    sDay = Worksheets(slotsSheet).Cells(sRow, dayCol).Value
    sPoints = Worksheets(slotsSheet).Cells(sRow, pointsCol).Value
    
    If (DutySlots.getColHeader(sCol) = "ACTUAL (ARMED)") Then
        sArmed = True
    Else
        sArmed = False
    End If
    
    ' Pre-allocated
    If (Worksheets(slotsSheet).Cells(sRow, sCol).Value <> "") Then
        sLocked = True
        sPersonnel = Worksheets(slotsSheet).Cells(sRow, sCol).Value
    End If
End Sub

Sub setVolunteer(name As String)
    sPersonnel = name
    Worksheets(slotsSheet).Cells(sRow, sCol).Value = sPersonnel
    Worksheets(slotsSheet).Cells(sRow, sCol).Font.Bold = True
    sLocked = True
End Sub

Sub writeToDutyList()
    ' Assign the guards and stand-by guards
    Worksheets(slotsSheet).Cells(sRow, sCol).Value = sPersonnel
    Worksheets(slotsSheet).Cells(sRow, sStbCol).Value = sStandby
    
    Dim pers As DutyPersonnel
    If (sPersonnel <> "") Then
        Set pers = New DutyPersonnel
        pers.initialize (sPersonnel)
        
        If (Not pers.armed) Then
            Worksheets(slotsSheet).Cells(sRow, sCol).Font.ColorIndex = 3
        End If
    End If
    
    If (sStandby <> "") Then
        Set pers = New DutyPersonnel
        pers.initialize (sStandby)
        
        If (Not pers.armed) Then
            Worksheets(slotsSheet).Cells(sRow, sStbCol).Font.ColorIndex = 3
        End If
    End If
End Sub

Sub HighlightEmpty()

    Dim colo As Integer
    
    colo = 35

    If (sPersonnel = "") Then
        Worksheets(slotsSheet).Cells(sRow, sCol).Interior.ColorIndex = colo
    End If
    
    If (sStandby = "") Then
        Worksheets(slotsSheet).Cells(sRow, sStbCol).Interior.ColorIndex = colo
    End If
End Sub

Function toString()
    Dim output As String
    output = sRow & "," & sCol & " PTS:" & sPoints & " Date:" & sDay & " DIF:" & sDifficulty & " PER:" & sPersonnel & " STB:" & sStandby
    
    toString = output
End Function

Public Property Get day() As String
    day = sDay
End Property

Public Property Get difficulty() As String
    difficulty = sDifficulty
End Property

Public Property Let difficulty(d As String)
    sDifficulty = d
End Property

Public Property Get personnel() As String
    personnel = sPersonnel
End Property

Public Property Let personnel(p As String)
    sPersonnel = p
End Property

Public Property Get standby() As String
    standby = sStandby
End Property

Public Property Let standby(p As String)
    sStandby = p
End Property

Public Property Get locked() As Boolean
    locked = sLocked
End Property

Public Property Let locked(l As Boolean)
    sLocked = l
End Property

Public Property Get points() As String
    points = sPoints
End Property

Public Property Get armed() As Boolean
    armed = sArmed
End Property


