Attribute VB_Name = "modMain"
Option Explicit

Public Sub ShowStartScreen()
    On Error GoTo ErrorHandler
    modUI.OnWorkbookOpen
    Exit Sub
ErrorHandler:
    If Application.DisplayAlerts Then
        MsgBox "Unable to open the START screen: " & Err.Description, vbExclamation, "Mega Pipes"
    End If
End Sub
