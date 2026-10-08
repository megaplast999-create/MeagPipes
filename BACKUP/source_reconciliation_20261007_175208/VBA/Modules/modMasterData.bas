Attribute VB_Name = "modMasterData"
Option Explicit

Public Sub SetupPlantDropdown()
    ApplyListValidation ThisWorkbook.Worksheets("QUOTE").Range("B14"), _
        TableColumnAddress("PLANTS", "tblPlants", "PlantName")
End Sub

Public Sub SetupPaymentTermsDropdown()
    ApplyListValidation ThisWorkbook.Worksheets("QUOTE").Range("B15"), _
        TableColumnAddress("MASTERS", "tblPaymentTerms", "PaymentTerm")
End Sub

Public Sub SetupPriceBasisDropdown()
    With ThisWorkbook.Worksheets("QUOTE").Range("B17").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Metres,MT,Nos"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Invalid price basis"
        .ErrorMessage = "Select Metres, MT, or Nos."
    End With
End Sub

Public Sub SetupMasterDrivenHeader()
    On Error GoTo ErrorHandler
    ApplyListValidation ThisWorkbook.Worksheets("QUOTE").Range("B8"), _
        TableColumnAddress("CUSTOMERS", "tblCustomers", "CustomerName")
    SetupPlantDropdown
    SetupPaymentTermsDropdown
    SetupPriceBasisDropdown
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "SetupMasterDrivenHeader", Err.Description
End Sub

Public Sub RefreshCustomerFields()
    Dim customerName As String
    Dim customerRow As ListRow
    customerName = Trim$(CStr(ThisWorkbook.Worksheets("QUOTE").Range("B8").Value))
    If Len(customerName) = 0 Then Exit Sub

    Set customerRow = FindTableRow("CUSTOMERS", "tblCustomers", "CustomerName", customerName)
    If customerRow Is Nothing Then
        Err.Raise vbObjectError + 1201, "RefreshCustomerFields", _
            "Selected customer was not found in tblCustomers."
    End If

    With customerRow.Range
        ThisWorkbook.Worksheets("QUOTE").Range("B9").Value = .Cells(1, TableColumnIndex("tblCustomers", "ContactPerson")).Value
        ThisWorkbook.Worksheets("QUOTE").Range("B11").Value = .Cells(1, TableColumnIndex("tblCustomers", "Address")).Value
    End With
End Sub

Private Sub ApplyListValidation(ByVal target As Range, ByVal sourceAddress As String)
    If Len(sourceAddress) = 0 Then
        Err.Raise vbObjectError + 1202, "ApplyListValidation", _
            "The master table has no usable values."
    End If
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=" & sourceAddress
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With
End Sub

Private Function TableColumnAddress(ByVal sheetName As String, ByVal tableName As String, ByVal columnName As String) As String
    Dim tableObject As ListObject
    Dim dataRange As Range
    Set tableObject = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If tableObject.DataBodyRange Is Nothing Then Exit Function
    Set dataRange = tableObject.ListColumns(columnName).DataBodyRange
    TableColumnAddress = "'" & sheetName & "'!" & dataRange.Address
End Function

Private Function FindTableRow(ByVal sheetName As String, ByVal tableName As String, ByVal columnName As String, ByVal expectedValue As String) As ListRow
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Dim columnNumber As Long
    Set tableObject = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If tableObject.DataBodyRange Is Nothing Then Exit Function
    columnNumber = tableObject.ListColumns(columnName).Index
    For Each rowObject In tableObject.ListRows
        If StrComp(Trim$(CStr(rowObject.Range.Cells(1, columnNumber).Value)), expectedValue, vbTextCompare) = 0 Then
            Set FindTableRow = rowObject
            Exit Function
        End If
    Next rowObject
End Function

Private Function TableColumnIndex(ByVal tableName As String, ByVal columnName As String) As Long
    TableColumnIndex = ThisWorkbook.Worksheets("CUSTOMERS").ListObjects(tableName).ListColumns(columnName).Index
End Function
