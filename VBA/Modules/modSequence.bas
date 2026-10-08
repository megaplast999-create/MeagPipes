Attribute VB_Name = "modSequence"
Option Explicit

Public Function AllocateEnquiryNumber() As String
    Dim settings As ListObject
    Dim serialRow As ListRow
    Dim currentSerial As Long

    Set settings = ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")
    Set serialRow = FindSettingRow(settings, "CurrentEnquirySerial")
    If serialRow Is Nothing Then
        currentSerial = CLng(GetSettingValue(settings, "EnquirySerialStart", 595))
        Set serialRow = settings.ListRows.Add
        serialRow.Range.Cells(1, settings.ListColumns("SettingKey").Index).Value = "CurrentEnquirySerial"
        serialRow.Range.Cells(1, settings.ListColumns("DataType").Index).Value = "Integer"
        serialRow.Range.Cells(1, settings.ListColumns("Description").Index).Value = "DEV local sequence counter"
        serialRow.Range.Cells(1, settings.ListColumns("Editable").Index).Value = "No"
        serialRow.Range.Cells(1, settings.ListColumns("Category").Index).Value = "Sequence"
    Else
        currentSerial = CLng(serialRow.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value)
    End If

    serialRow.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value = currentSerial + 1
    AllocateEnquiryNumber = "MKTG/ENQ/26-27/" & CStr(currentSerial)
End Function

Private Function FindSettingRow(ByVal settings As ListObject, ByVal key As String) As ListRow
    Dim rowObject As ListRow
    For Each rowObject In settings.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, settings.ListColumns("SettingKey").Index).Value), key, vbTextCompare) = 0 Then
            Set FindSettingRow = rowObject
            Exit Function
        End If
    Next rowObject
End Function

Private Function GetSettingValue(ByVal settings As ListObject, ByVal key As String, ByVal fallback As Long) As Long
    Dim rowObject As ListRow
    Set rowObject = FindSettingRow(settings, key)
    If rowObject Is Nothing Then
        GetSettingValue = fallback
    ElseIf Len(Trim$(CStr(rowObject.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value))) = 0 Then
        GetSettingValue = fallback
    Else
        GetSettingValue = CLng(rowObject.Range.Cells(1, settings.ListColumns("SettingValue").Index).Value)
    End If
End Function
