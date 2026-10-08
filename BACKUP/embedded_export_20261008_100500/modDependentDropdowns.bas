Attribute VB_Name = "modDependentDropdowns"
Option Explicit

Private Const HELPER_FIRST_ROW As Long = 2

Public Sub SetupDependentDropdownArchitecture()
    On Error GoTo ErrorHandler
    BuildHelperList "ProductList", "tblProducts", "ProductName", "", 27
    BuildHelperList "FamilyList", "tblProducts", "Family", "", 28
    BuildHelperList "BasisList", "tblSizes", "Basis", "", 29
    BuildHelperList "ExternalCoatingList", "tblExternalCoatings", "CoatingName", "", 30
    BuildHelperList "InternalCoatingList", "tblInternalCoatings", "CoatingName", "", 31
    ApplyNamedListValidation ThisWorkbook.Worksheets("QUOTE").Range("B21"), "ProductList"
    ApplyNamedListValidation ThisWorkbook.Worksheets("QUOTE").Range("B22"), "FamilyList"
    ApplyNamedListValidation ThisWorkbook.Worksheets("QUOTE").Range("B23"), "BasisList"
    ApplyNamedListValidation ThisWorkbook.Worksheets("QUOTE").Range("B27"), "ExternalCoatingList"
    ApplyNamedListValidation ThisWorkbook.Worksheets("QUOTE").Range("B28"), "InternalCoatingList"
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "SetupDependentDropdownArchitecture", Err.Description
End Sub

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
    masterSheet.Cells(1, helperColumn).value = listName
    outputRow = HELPER_FIRST_ROW
    For Each rowObject In sourceTable.ListRows
        value = Trim$(CStr(rowObject.Range.Cells(1, sourceTable.ListColumns(columnName).Index).value))
        If Len(value) > 0 And (Len(filterValue) = 0 Or StrComp(value, filterValue, vbTextCompare) = 0) Then
            masterSheet.Cells(outputRow, helperColumn).value = value
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

Private Sub ApplyNamedListValidation(ByVal target As Range, ByVal listName As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With
End Sub

Private Function FindTable(ByVal tableName As String) As ListObject
    Dim sheetObject As Worksheet
    For Each sheetObject In ThisWorkbook.Worksheets
        On Error Resume Next
        Set FindTable = sheetObject.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTable Is Nothing Then Exit Function
    Next sheetObject
End Function
