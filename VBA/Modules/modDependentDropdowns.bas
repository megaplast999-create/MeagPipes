Attribute VB_Name = "modDependentDropdowns"
Option Explicit

Private Const HELPER_FIRST_ROW As Long = 2

Public Sub SetupDependentDropdownArchitecture()
    Dim wsM As Worksheet
    Dim quoteSheet As Worksheet
    On Error GoTo ErrorHandler
    Set wsM = ThisWorkbook.Worksheets("MASTERS")
    Set quoteSheet = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet wsM
    modProtection.UnprotectSheet quoteSheet
    
    BuildHelperList "ProductList", "tblProducts", "ProductName", "", 27
    BuildHelperList "FamilyList", "tblProducts", "Family", "", 28
    BuildHelperList "BasisList", "tblSizes", "Basis", "", 29
    BuildHelperList "ExternalCoatingList", "tblExternalCoatings", "CoatingName", "", 30
    BuildHelperList "InternalCoatingList", "tblInternalCoatings", "CoatingName", "", 31
    ApplyNamedListValidation quoteSheet.Range("B20"), "ProductList"
    ApplyNamedListValidation quoteSheet.Range("B21"), "FamilyList"
    ApplyNamedListValidation quoteSheet.Range("B22"), "BasisList"
    ApplyNamedListValidation quoteSheet.Range("B26"), "ExternalCoatingList"
    ApplyNamedListValidation quoteSheet.Range("B27"), "InternalCoatingList"
    SetupQuoteLineDropdowns
    modProtection.ProtectSheet quoteSheet, userInterfaceOnly:=True
    modProtection.ProtectSheet wsM, userInterfaceOnly:=True
    Exit Sub
ErrorHandler:
    On Error Resume Next
    modProtection.ProtectSheet quoteSheet, userInterfaceOnly:=True
    modProtection.ProtectSheet wsM, userInterfaceOnly:=True
    On Error GoTo 0
    Err.Raise Err.Number, "SetupDependentDropdownArchitecture", Err.Description
End Sub

Public Sub SetupQuoteLineDropdowns()
    Dim quoteSheet As Worksheet
    Dim wsM As Worksheet
    Dim family As String
    Dim basis As String

    On Error GoTo ErrorHandler
    Set quoteSheet = ThisWorkbook.Worksheets("QUOTE")
    Set wsM = ThisWorkbook.Worksheets("MASTERS")
    modProtection.UnprotectSheet wsM
    modProtection.UnprotectSheet quoteSheet

    family = Trim$(CStr(quoteSheet.Range("B21").Value2))
    basis = UCase$(Trim$(CStr(quoteSheet.Range("B22").Value2)))
    If Len(family) = 0 Then family = "HSAW"
    If Len(basis) = 0 Then basis = "ID"

    If Len(family) > 0 And (basis = "ID" Or basis = "OD") Then
        BuildSizeHelperList family, basis, 32
        ApplyNamedListValidation quoteSheet.Range("B23"), "Sizes_" & family & "_" & basis
    End If

    If Len(family) > 0 Then
        BuildFilteredHelperList "WT_" & family, "tblThickness", "Family", "ThicknessMM", family, 33
        BuildFilteredHelperList "Grades_" & family, "tblGrades", "Family", "GradeName", family, 34
        ApplyNamedListValidation quoteSheet.Range("B24"), "WT_" & family
        ApplyNamedListValidation quoteSheet.Range("B25"), "Grades_" & family
    End If
    ApplyPositiveValidation quoteSheet.Range("B28")
    ApplyInlineValidation quoteSheet.Range("B29"), "Metres,MT,Nos", "Select Metres, MT, or Nos."
    modProtection.ProtectSheet quoteSheet, userInterfaceOnly:=True
    modProtection.ProtectSheet wsM, userInterfaceOnly:=True
    Exit Sub
ErrorHandler:
    On Error Resume Next
    modProtection.ProtectSheet quoteSheet, userInterfaceOnly:=True
    modProtection.ProtectSheet wsM, userInterfaceOnly:=True
    On Error GoTo 0
    Err.Raise Err.Number, "SetupQuoteLineDropdowns", Err.Description
End Sub

Public Sub RequestNewSize(ByVal family As String, ByVal basis As String, _
                          ByVal sizeMm As Double, Optional ByVal notes As String = "")
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Dim requestFolder As String
    Dim requestPath As String
    Dim fileNumber As Integer

    family = UCase$(Trim$(family))
    basis = UCase$(Trim$(basis))
    ValidateNewSize family, basis, sizeMm
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblSizes")
    If SizeExists(tableObject, family, basis, sizeMm) Then
        Err.Raise vbObjectError + 1510, "RequestNewSize", "This size already exists."
    End If

    Set rowObject = tableObject.ListRows.Add
    With rowObject.Range
        .Cells(1, tableObject.ListColumns("SizeID").Index).Value = _
            family & "-" & basis & "-" & Format$(tableObject.ListRows.Count, "000")
        .Cells(1, tableObject.ListColumns("Family").Index).Value = family
        .Cells(1, tableObject.ListColumns("Basis").Index).Value = basis
        .Cells(1, tableObject.ListColumns("SizeMM").Index).Value = sizeMm
        .Cells(1, tableObject.ListColumns("Active").Index).Value = "Yes"
        .Cells(1, tableObject.ListColumns("Pending").Index).Value = "Yes"
        .Cells(1, tableObject.ListColumns("Notes").Index).Value = notes
    End With

    requestFolder = ThisWorkbook.Path & Application.PathSeparator & "Requests"
    If Dir$(requestFolder, vbDirectory) = vbNullString Then MkDir requestFolder
    requestPath = requestFolder & Application.PathSeparator & _
        "size_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Environ$("USERNAME") & ".csv"
    fileNumber = FreeFile
    Open requestPath For Output As #fileNumber
    Print #fileNumber, "Family,Basis,SizeMM,Notes,User,RequestedAt"
    Print #fileNumber, CsvValue(family) & "," & CsvValue(basis) & "," & _
        CStr(sizeMm) & "," & CsvValue(notes) & "," & CsvValue(Environ$("USERNAME")) & "," & _
        Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #fileNumber
    SetupQuoteLineDropdowns
End Sub

Public Function IsSizeWithinRange(ByVal family As String, ByVal basis As String, _
                                  ByVal sizeMm As Double) As Boolean
    On Error GoTo InvalidSize
    ValidateNewSize UCase$(Trim$(family)), UCase$(Trim$(basis)), sizeMm
    IsSizeWithinRange = True
    Exit Function
InvalidSize:
    IsSizeWithinRange = False
End Function

Private Sub BuildHelperList(ByVal listName As String, ByVal tableName As String, ByVal columnName As String, ByVal filterValue As String, ByVal helperColumn As Long)
    Dim masterSheet As Worksheet
    Dim sourceTable As ListObject
    Dim rowObject As ListRow
    Dim outputRow As Long
    Dim value As String
    Dim helperRange As Range

    Set masterSheet = ThisWorkbook.Worksheets("MASTERS")
    Set sourceTable = FindTable(tableName)
    If sourceTable Is Nothing Then Err.Raise vbObjectError + 1501, "BuildHelperList", "Master table not found: " & tableName

    masterSheet.Columns(helperColumn).ClearContents
    masterSheet.Cells(1, helperColumn).Value = listName
    outputRow = HELPER_FIRST_ROW
    For Each rowObject In sourceTable.ListRows
        value = Trim$(CStr(rowObject.Range.Cells(1, sourceTable.ListColumns(columnName).Index).Value))
        If Len(value) > 0 And (Len(filterValue) = 0 Or StrComp(value, filterValue, vbTextCompare) = 0) _
            And Not HelperValueExists(masterSheet, helperColumn, outputRow - 1, value) Then
            masterSheet.Cells(outputRow, helperColumn).Value = value
            outputRow = outputRow + 1
        End If
    Next rowObject
    If outputRow = HELPER_FIRST_ROW Then Err.Raise vbObjectError + 1502, "BuildHelperList", "No values found for " & listName

    Set helperRange = masterSheet.Range(masterSheet.Cells(HELPER_FIRST_ROW, helperColumn), masterSheet.Cells(outputRow - 1, helperColumn))
    On Error Resume Next
    ThisWorkbook.Names(listName).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=listName, RefersTo:="=" & helperRange.Address(External:=True)
End Sub

Private Function HelperValueExists(ByVal sheetObject As Worksheet, ByVal helperColumn As Long, _
                                  ByVal lastRow As Long, ByVal value As String) As Boolean
    Dim rowNumber As Long
    If lastRow < HELPER_FIRST_ROW Then Exit Function
    For rowNumber = HELPER_FIRST_ROW To lastRow
        If StrComp(CStr(sheetObject.Cells(rowNumber, helperColumn).Value), value, vbTextCompare) = 0 Then
            HelperValueExists = True
            Exit Function
        End If
    Next rowNumber
End Function

Private Sub BuildSizeHelperList(ByVal family As String, ByVal basis As String, ByVal helperColumn As Long)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Dim outputRow As Long
    Dim listName As String
    Set tableObject = FindTable("tblSizes")
    If tableObject Is Nothing Then Err.Raise vbObjectError + 1503, "BuildSizeHelperList", "Size master table not found."
    listName = "Sizes_" & family & "_" & basis
    ThisWorkbook.Worksheets("MASTERS").Columns(helperColumn).ClearContents
    ThisWorkbook.Worksheets("MASTERS").Cells(1, helperColumn).Value = listName
    outputRow = HELPER_FIRST_ROW
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").Index).Value), family, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Basis").Index).Value), basis, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").Index).Value), "Yes", vbTextCompare) = 0 Then
            ThisWorkbook.Worksheets("MASTERS").Cells(outputRow, helperColumn).Value = _
                rowObject.Range.Cells(1, tableObject.ListColumns("SizeMM").Index).Value
            outputRow = outputRow + 1
        End If
    Next rowObject
    If outputRow = HELPER_FIRST_ROW Then Err.Raise vbObjectError + 1504, "BuildSizeHelperList", "No sizes found."
    CreateHelperName listName, helperColumn, outputRow - 1
End Sub

Private Sub BuildFilteredHelperList(ByVal listName As String, ByVal tableName As String, _
                                    ByVal filterColumn As String, ByVal valueColumn As String, _
                                    ByVal filterValue As String, ByVal helperColumn As Long)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Dim outputRow As Long
    Set tableObject = FindTable(tableName)
    If tableObject Is Nothing Then Err.Raise vbObjectError + 1505, "BuildFilteredHelperList", "Master table not found: " & tableName
    ThisWorkbook.Worksheets("MASTERS").Columns(helperColumn).ClearContents
    ThisWorkbook.Worksheets("MASTERS").Cells(1, helperColumn).Value = listName
    outputRow = HELPER_FIRST_ROW
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns(filterColumn).Index).Value), filterValue, vbTextCompare) = 0 Then
            ThisWorkbook.Worksheets("MASTERS").Cells(outputRow, helperColumn).Value = _
                rowObject.Range.Cells(1, tableObject.ListColumns(valueColumn).Index).Value
            outputRow = outputRow + 1
        End If
    Next rowObject
    If outputRow = HELPER_FIRST_ROW Then Err.Raise vbObjectError + 1506, "BuildFilteredHelperList", "No values found for " & listName
    CreateHelperName listName, helperColumn, outputRow - 1
End Sub

Private Sub CreateHelperName(ByVal listName As String, ByVal helperColumn As Long, ByVal lastRow As Long)
    Dim helperRange As Range
    On Error Resume Next
    ThisWorkbook.Names(listName).Delete
    On Error GoTo 0
    Set helperRange = ThisWorkbook.Worksheets("MASTERS").Range( _
        ThisWorkbook.Worksheets("MASTERS").Cells(HELPER_FIRST_ROW, helperColumn), _
        ThisWorkbook.Worksheets("MASTERS").Cells(lastRow, helperColumn))
    ThisWorkbook.Names.Add Name:=listName, RefersTo:="=" & helperRange.Address(External:=True)
End Sub

Private Sub ApplyNamedListValidation(ByVal target As Range, ByVal listName As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With
End Sub

Private Sub ApplyInlineValidation(ByVal target As Range, ByVal values As String, ByVal errorMessage As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=values
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Invalid value"
        .ErrorMessage = errorMessage
    End With
End Sub

Private Sub ApplyPositiveValidation(ByVal target As Range)
    With target.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
            Operator:=xlGreater, Formula1:="0"
        .IgnoreBlank = True
        .ShowError = True
        .ErrorTitle = "Invalid quantity"
        .ErrorMessage = "Quantity must be greater than zero."
    End With
End Sub

Private Sub ValidateNewSize(ByVal family As String, ByVal basis As String, ByVal sizeMm As Double)
    Dim minimum As Double
    Dim maximum As Double
    If (family <> "HSAW" And family <> "HFW") Or (basis <> "ID" And basis <> "OD") Then
        Err.Raise vbObjectError + 1507, "ValidateNewSize", "Family must be HSAW or HFW and basis must be ID or OD."
    End If
    If family = "HSAW" Then
        minimum = 457
        maximum = 2540
    Else
        minimum = 168.3
        maximum = 508
    End If
    If basis = "ID" Then
        maximum = maximum - 2# * FamilyMaxThickness(family)
        minimum = minimum - 2# * FamilyMaxThickness(family)
    End If
    If sizeMm < minimum Or sizeMm > maximum Then
        Err.Raise vbObjectError + 1508, "ValidateNewSize", "Size is outside the allowed family range."
    End If
End Sub

Private Function FamilyMaxThickness(ByVal family As String) As Double
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = FindTable("tblThickness")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").Index).Value), family, vbTextCompare) = 0 Then
            If CDbl(rowObject.Range.Cells(1, tableObject.ListColumns("MaxMM").Index).Value) > FamilyMaxThickness Then
                FamilyMaxThickness = CDbl(rowObject.Range.Cells(1, tableObject.ListColumns("MaxMM").Index).Value)
            End If
        End If
    Next rowObject
    If FamilyMaxThickness <= 0 Then Err.Raise vbObjectError + 1509, "FamilyMaxThickness", "No thickness range found."
End Function

Private Function SizeExists(ByVal tableObject As ListObject, ByVal family As String, _
                            ByVal basis As String, ByVal sizeMm As Double) As Boolean
    Dim rowObject As ListRow
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").Index).Value), family, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Basis").Index).Value), basis, vbTextCompare) = 0 _
            And Abs(CDbl(rowObject.Range.Cells(1, tableObject.ListColumns("SizeMM").Index).Value) - sizeMm) < 0.0001 Then
            SizeExists = True
            Exit Function
        End If
    Next rowObject
End Function

Private Function CsvValue(ByVal value As String) As String
    CsvValue = """" & Replace$(value, """", """""") & """"
End Function

Private Function FindTable(ByVal tableName As String) As ListObject
    Dim sheetObject As Worksheet
    For Each sheetObject In ThisWorkbook.Worksheets
        On Error Resume Next
        Set FindTable = sheetObject.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTable Is Nothing Then Exit Function
    Next sheetObject
End Function
