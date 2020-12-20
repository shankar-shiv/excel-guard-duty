Rem Attribute VBA_ModuleType=VBAClassModule

Option Explicit

Private logSheet As String
Private logCell As String

Private Sub Class_Initialize()
    logSheet = "Guard Duty Planner"
    logCell = "B7"
    ' Clear log
    Me.clearLog
    
End Sub

Public Function log(message As String)
    ' VBcrlf = Visual Basic Carriage return line feed
    Worksheets(logSheet).Range(logCell).Value = Worksheets(logSheet).Range(logCell).Value & message & vbCrLf
End Function


Public Function clearLog()
    Worksheets(logSheet).Range(logCell).Value = ""
End Function
