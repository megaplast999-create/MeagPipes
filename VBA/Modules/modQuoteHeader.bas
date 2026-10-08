Attribute VB_Name = "modQuoteHeader"
Option Explicit

Public Sub InitializeNewQuote()
    On Error GoTo ErrorHandler
    With ThisWorkbook.Worksheets("QUOTE")
        .Range("B4").Value = AllocateEnquiryNumber()
        .Range("B5").Value = Date
        .Range("B16").Value = GetConfiguredValidity()
        .Range("B17").Value = "Metres"
    End With
    SetupMasterDrivenHeader
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "InitializeNewQuote", Err.Description
End Sub

Public Sub HandleCustomerSelection()
    On Error GoTo ErrorHandler
    RefreshCustomerFields
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "HandleCustomerSelection", Err.Description
End Sub

Public Function GetConfiguredValidity() As Long
    Dim settings As ListObject
    Dim rowObject As ListRow
    Set settings = ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")
    For Each rowObject In settings.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, settings.ListColumns("SettingKey").Index).Value), "ValidityDays", vbTextCompare) = 0 Then
            GetConfiguredValidity = CLng(rowObject.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value)
            Exit Function
        End If
    Next rowObject
    Err.Raise vbObjectError + 1401, "GetConfiguredValidity", "ValidityDays is missing from tblSettings."
End Function
