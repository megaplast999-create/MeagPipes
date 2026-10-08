VERSION 1.0 CLASS
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ShowStartScreen

    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:

    Application.ScreenUpdating = True

    MsgBox "Mega Pipes Quotation System could not open the START screen." & _
           vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbExclamation, _
           "Mega Pipes Quotation System"

End Sub
