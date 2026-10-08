Attribute VB_Name = "modMain"
Option Explicit

Public Sub ShowStartScreen()
    On Error GoTo ErrorHandler
    ThisWorkbook.Worksheets("START").Activate
    Exit Sub
ErrorHandler:
    MsgBox "Unable to open the START sheet: " & Err.Description, vbExclamation, "Mega Pipes"
End Sub
