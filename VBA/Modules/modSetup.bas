Attribute VB_Name = "modSetup"
Option Explicit

Public Sub SetupCustomerDropdown()

    Dim wsQuote As Worksheet
    Dim wsCustomers As Worksheet
    Dim lastRow As Long
    Dim customerRange As String

    On Error GoTo ErrorHandler

    Set wsQuote = ThisWorkbook.Worksheets("QUOTE")
    Set wsCustomers = ThisWorkbook.Worksheets("CUSTOMERS")

    'Find the last customer in CustomerName column B
    lastRow = wsCustomers.Cells(wsCustomers.Rows.Count, "B").End(xlUp).Row

    If lastRow < 4 Then
        MsgBox "No customers were found in the customer master.", _
               vbExclamation, _
               "Mega Pipes"
        Exit Sub
    End If

    'Remove previous validation
    With wsQuote.Range("B8")
        .Validation.Delete

        'Create customer dropdown
        customerRange = "='" & wsCustomers.Name & "'!$B$4:$B$" & lastRow

        .Validation.Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:=customerRange

        .Validation.IgnoreBlank = True
        .Validation.InCellDropdown = True
    End With

    MsgBox "Customer dropdown created successfully." & vbCrLf & vbCrLf & _
           "Customers found: " & (lastRow - 3), _
           vbInformation, _
           "Mega Pipes"

    Exit Sub

ErrorHandler:

    MsgBox "Unable to create the customer dropdown." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, _
           "Mega Pipes"

End Sub
