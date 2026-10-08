Attribute VB_Name = "modQuoteLines"
Option Explicit

Public Sub AddQuoteLine(ByVal quoteId As String, ByVal productId As String, ByVal family As String, _
    ByVal basis As String, ByVal sizeMm As Double, ByVal thicknessMm As Double, _
    ByVal gradeId As String, ByVal quantity As Double, ByVal quantityUnit As String, _
    ByVal pipeLengthM As Double, ByVal plantId As String)

    Dim tableObject As ListObject
    Dim lineObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")

    If Len(Trim$(quoteId)) = 0 Or Len(Trim$(productId)) = 0 Or Len(Trim$(family)) = 0 Or Len(Trim$(gradeId)) = 0 Then
        Err.Raise vbObjectError + 1301, "AddQuoteLine", "Quote, product, family, and grade are required."
    End If
    If Not IsValidQuantity(quantity) Then
        Err.Raise vbObjectError + 1302, "AddQuoteLine", "Quantity must be greater than zero."
    End If
    If Not IsValidPipeLength(pipeLengthM) Then
        Err.Raise vbObjectError + 1303, "AddQuoteLine", "Pipe length must be between 6 and 13 metres."
    End If
    If UCase$(Trim$(basis)) <> "OD" And UCase$(Trim$(basis)) <> "ID" Then
        Err.Raise vbObjectError + 1304, "AddQuoteLine", "Basis must be OD or ID."
    End If
    ValidateProductFamily productId, family
    ValidateSize family, basis, sizeMm
    ValidateThickness family, thicknessMm
    ValidateGrade family, gradeId
    ValidateQuantityUnit quantityUnit
    ValidatePlant plantId

    Set lineObject = tableObject.ListRows.Add
    With lineObject.Range
        .Cells(1, tableObject.ListColumns("LineID").index).value = "L" & Format$(tableObject.ListRows.Count, "0000")
        .Cells(1, tableObject.ListColumns("QuoteID").index).value = quoteId
        .Cells(1, tableObject.ListColumns("ProductID").index).value = productId
        .Cells(1, tableObject.ListColumns("Family").index).value = family
        .Cells(1, tableObject.ListColumns("Basis").index).value = basis
        .Cells(1, tableObject.ListColumns("SizeMM").index).value = sizeMm
        .Cells(1, tableObject.ListColumns("ThicknessMM").index).value = thicknessMm
        .Cells(1, tableObject.ListColumns("GradeID").index).value = gradeId
        .Cells(1, tableObject.ListColumns("Quantity").index).value = quantity
        .Cells(1, tableObject.ListColumns("QuantityUnit").index).value = quantityUnit
        .Cells(1, tableObject.ListColumns("PipeLengthM").index).value = pipeLengthM
        .Cells(1, tableObject.ListColumns("PlantID").index).value = plantId
        .Cells(1, tableObject.ListColumns("Status").index).value = "DRAFT"
    End With
End Sub

Private Sub ValidateProductFamily(ByVal productId As String, ByVal family As String)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblProducts")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("ProductID").index).value), productId, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").index).value), family, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").index).value), "Yes", vbTextCompare) = 0 Then Exit Sub
    Next rowObject
    Err.Raise vbObjectError + 1305, "ValidateProductFamily", "Product and family do not match an active master row."
End Sub

Private Sub ValidateSize(ByVal family As String, ByVal basis As String, ByVal sizeMm As Double)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblSizes")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").index).value), family, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Basis").index).value), basis, vbTextCompare) = 0 _
            And Abs(CDbl(rowObject.Range.Cells(1, tableObject.ListColumns("SizeMM").index).value) - sizeMm) < 0.0001 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").index).value), "Yes", vbTextCompare) = 0 Then Exit Sub
    Next rowObject
    Err.Raise vbObjectError + 1306, "ValidateSize", "Size is not available for the selected family and basis."
End Sub

Private Sub ValidateThickness(ByVal family As String, ByVal thicknessMm As Double)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblThickness")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").index).value), family, vbTextCompare) = 0 _
            And Abs(CDbl(rowObject.Range.Cells(1, tableObject.ListColumns("ThicknessMM").index).value) - thicknessMm) < 0.0001 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").index).value), "Yes", vbTextCompare) = 0 Then Exit Sub
    Next rowObject
    Err.Raise vbObjectError + 1307, "ValidateThickness", "Thickness is not available for the selected family."
End Sub

Private Sub ValidateGrade(ByVal family As String, ByVal gradeId As String)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblGrades")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").index).value), family, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("GradeID").index).value), gradeId, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").index).value), "Yes", vbTextCompare) = 0 Then Exit Sub
    Next rowObject
    Err.Raise vbObjectError + 1308, "ValidateGrade", "Grade is not available for the selected family."
End Sub

Private Sub ValidateQuantityUnit(ByVal quantityUnit As String)
    Select Case UCase$(Trim$(quantityUnit))
        Case "METRES", "MT", "NOS"
        Case Else
            Err.Raise vbObjectError + 1309, "ValidateQuantityUnit", "Unit must be Metres, MT, or Nos."
    End Select
End Sub

Private Sub ValidatePlant(ByVal plantId As String)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("PLANTS").ListObjects("tblPlants")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("PlantID").index).value), plantId, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").index).value), "Yes", vbTextCompare) = 0 Then Exit Sub
    Next rowObject
    Err.Raise vbObjectError + 1310, "ValidatePlant", "Plant is not available in the active plant master."
End Sub
