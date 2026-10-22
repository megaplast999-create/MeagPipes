Attribute VB_Name = "modMasterData"
Option Explicit

Public Sub SetupQuoteHeader()
    On Error GoTo ErrorHandler

    RequireTable "CUSTOMERS", "tblCustomers", "CustomerName"
    RequireTable "PLANTS", "tblPlants", "PlantName"
    RequireTable "MASTERS", "tblPaymentTerms", "PaymentTerm"
    RequireTable "SETTINGS", "tblSettings", "SettingKey"

    CreateTableColumnName "CustomerNames", "CUSTOMERS", "tblCustomers", "CustomerName"
    CreateTableColumnName "PlantNames", "PLANTS", "tblPlants", "PlantName"
    CreateTableColumnName "PaymentTermDescriptions", "MASTERS", "tblPaymentTerms", "PaymentTerm"

    ApplyNamedValidation ThisWorkbook.Worksheets("QUOTE").Range("B8"), "CustomerNames", _
        "Select a customer from tblCustomers."
    ApplyNamedValidation ThisWorkbook.Worksheets("QUOTE").Range("B13"), "PlantNames", _
        "Select a plant from tblPlants."
    ApplyNamedValidation ThisWorkbook.Worksheets("QUOTE").Range("B14"), "PaymentTermDescriptions", _
        "Select a payment term from tblPaymentTerms."
    ApplyInlineValidation ThisWorkbook.Worksheets("QUOTE").Range("B16"), _
        "Metres,MT,Nos", "Select Metres, MT, or Nos."
    ApplyInlineValidation ThisWorkbook.Worksheets("QUOTE").Range("B17"), _
        "NORMAL,LEGACY", "Select Pricing Mode."
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "SetupQuoteHeader", Err.Description
End Sub

Public Sub SetupPlantDropdown()
    ApplyListValidation ThisWorkbook.Worksheets("QUOTE").Range("B13"), _
        TableColumnAddress("PLANTS", "tblPlants", "PlantName")
End Sub

Private Sub RequireTable(ByVal sheetName As String, ByVal tableName As String, ByVal columnName As String)
    Dim tableObject As ListObject
    Set tableObject = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If tableObject.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1200, "RequireTable", _
            "Required master table has no data rows: " & tableName
    End If
    If tableObject.ListColumns(columnName).DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1201, "RequireTable", _
            "Required master column has no data: " & tableName & "." & columnName
    End If
End Sub

Private Sub CreateTableColumnName(ByVal nameText As String, ByVal sheetName As String, ByVal tableName As String, ByVal columnName As String)
    Dim dataRange As Range
    Set dataRange = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName).ListColumns(columnName).DataBodyRange
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nameText, RefersTo:="=" & dataRange.Address(External:=True)
End Sub

Private Sub ApplyNamedValidation(ByVal target As Range, ByVal nameText As String, ByVal errorMessage As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=" & nameText
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Invalid selection"
        .ErrorMessage = errorMessage
    End With
End Sub

Private Sub ApplyInlineValidation(ByVal target As Range, ByVal values As String, ByVal errorMessage As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=values
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Invalid selection"
        .ErrorMessage = errorMessage
    End With
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
    SetupQuoteHeader
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
        ThisWorkbook.Worksheets("QUOTE").Range("B10").Value = .Cells(1, TableColumnIndex("tblCustomers", "Address")).Value
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
    Dim cellVal As String
    
    Set tableObject = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If tableObject.DataBodyRange Is Nothing Then Exit Function
    columnNumber = tableObject.ListColumns(columnName).Index
    
    ' Pass 1: Exact match
    For Each rowObject In tableObject.ListRows
        cellVal = Trim$(CStr(rowObject.Range.Cells(1, columnNumber).Value))
        If StrComp(cellVal, expectedValue, vbTextCompare) = 0 Then
            Set FindTableRow = rowObject
            Exit Function
        End If
    Next rowObject
    
    ' Pass 2: Substring or prefix match
    For Each rowObject In tableObject.ListRows
        cellVal = Trim$(CStr(rowObject.Range.Cells(1, columnNumber).Value))
        If InStr(1, cellVal, expectedValue, vbTextCompare) > 0 Or InStr(1, expectedValue, cellVal, vbTextCompare) > 0 Then
            Set FindTableRow = rowObject
            Exit Function
        End If
    Next rowObject
    
    ' Pass 3: Check CustomerID column if present
    On Error Resume Next
    Dim idCol As Long
    idCol = tableObject.ListColumns("CustomerID").Index
    If Err.Number = 0 Then
        For Each rowObject In tableObject.ListRows
            cellVal = Trim$(CStr(rowObject.Range.Cells(1, idCol).Value))
            If StrComp(cellVal, expectedValue, vbTextCompare) = 0 Then
                Set FindTableRow = rowObject
                Exit Function
            End If
        Next rowObject
    End If
    On Error GoTo 0
End Function

Private Function TableColumnIndex(ByVal tableName As String, ByVal columnName As String) As Long
    TableColumnIndex = ThisWorkbook.Worksheets("CUSTOMERS").ListObjects(tableName).ListColumns(columnName).Index
End Function
