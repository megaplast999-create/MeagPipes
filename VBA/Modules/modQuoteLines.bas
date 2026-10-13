Attribute VB_Name = "modQuoteLines"
Option Explicit

Public Sub AddQuoteLine(ByVal quoteId As String, ByVal productId As String, ByVal family As String, _
    ByVal basis As String, ByVal sizeMm As Double, ByVal thicknessMm As Double, _
    ByVal gradeId As String, ByVal quantity As Double, ByVal quantityUnit As String, _
    ByVal pipeLengthM As Double, ByVal plantId As String, _
    Optional ByVal legacyMode As Boolean = False, _
    Optional ByVal priceByMT As Boolean = False, _
    Optional ByVal manualFC As Variant, _
    Optional ByVal tripRate As Variant)

    Dim tableObject As ListObject
    Dim lineObject As ListRow
    Dim lineData As QuoteLineInput
    Dim rates As RateInput
    Dim result As LineResult
    Dim hasCoating As Boolean
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

    If tableObject.ListRows.Count = 1 And _
       Len(Trim$(CStr(tableObject.ListRows(1).Range.Cells(1, tableObject.ListColumns("ProductID").Index).Value))) = 0 Then
        Set lineObject = tableObject.ListRows(1)
    Else
        Set lineObject = tableObject.ListRows.Add
    End If
    With lineObject.Range
        .Cells(1, tableObject.ListColumns("LineID").Index).Value = "L" & Format$(lineObject.Index, "0000")
        .Cells(1, tableObject.ListColumns("QuoteID").Index).Value = quoteId
        .Cells(1, tableObject.ListColumns("ProductID").Index).Value = productId
        .Cells(1, tableObject.ListColumns("Family").Index).Value = family
        .Cells(1, tableObject.ListColumns("Basis").Index).Value = basis
        .Cells(1, tableObject.ListColumns("SizeMM").Index).Value = sizeMm
        .Cells(1, tableObject.ListColumns("ThicknessMM").Index).Value = thicknessMm
        .Cells(1, tableObject.ListColumns("GradeID").Index).Value = gradeId
        .Cells(1, tableObject.ListColumns("Quantity").Index).Value = quantity
        .Cells(1, tableObject.ListColumns("QuantityUnit").Index).Value = quantityUnit
        .Cells(1, tableObject.ListColumns("PipeLengthM").Index).Value = pipeLengthM
        .Cells(1, tableObject.ListColumns("PlantID").Index).Value = plantId
        .Cells(1, tableObject.ListColumns("Status").Index).Value = "DRAFT"
    End With

    On Error GoTo CalculationError

    hasCoating = (InStr(1, productId, "02", vbTextCompare) > 0 Or _
                  InStr(1, productId, "04", vbTextCompare) > 0)
    With lineData
        .Basis = basis
        .SizeMM = sizeMm
        .WT = thicknessMm
        .Qty = quantity
        .Unit = quantityUnit
        .PipeLen = pipeLengthM
        .PerTrailer = CDec(4)
        .HasExt = hasCoating
        .HasInt = hasCoating
    End With

    LoadDefaultRateInput rates, legacyMode
    rates.ExtRate = CDec(500)
    rates.IntRate = CDec(275)
    If IsMissing(tripRate) Then
        rates.TripRate = CDec(SettingValue("TripRateManual"))
    Else
        rates.TripRate = CDec(tripRate)
    End If
    If legacyMode Then
        If IsMissing(manualFC) Then
            rates.ManualFC = CDec(SettingValue("DefaultFinancialPct"))
        Else
            rates.ManualFC = CDec(manualFC)
        End If
    End If

    result = CalcLine(lineData, rates, priceByMT)
    WriteCalculatedQuoteLine lineObject, tableObject, result, priceByMT
    Exit Sub

CalculationError:
    On Error Resume Next
    lineObject.Delete
    On Error GoTo 0
    Err.Raise Err.Number, "AddQuoteLine", Err.Description
End Sub

Private Sub WriteCalculatedQuoteLine(ByVal lineObject As ListRow, _
                                     ByVal tableObject As ListObject, _
                                     ByRef result As LineResult, _
                                     ByVal priceByMT As Boolean)
    With lineObject.Range
        .Cells(1, tableObject.ListColumns("CalculatedODMM").Index).Value = result.od
        .Cells(1, tableObject.ListColumns("CalculatedIDMM").Index).Value = result.ID
        .Cells(1, tableObject.ListColumns("MassKgPerM").Index).Value = result.MassPerM
        .Cells(1, tableObject.ListColumns("Metres").Index).Value = result.Metres
        .Cells(1, tableObject.ListColumns("PipeCount").Index).Value = result.Pipes
        .Cells(1, tableObject.ListColumns("WeightMT").Index).Value = result.MT
        .Cells(1, tableObject.ListColumns("UnitPrice").Index).Value = _
            IIf(priceByMT, result.TotalPerMT, result.TotalPerM)
        .Cells(1, tableObject.ListColumns("PriceUnit").Index).Value = _
            IIf(priceByMT, "MT", "Metres")
        .Cells(1, tableObject.ListColumns("LineTotal").Index).Value = result.LineValue
        .Cells(1, tableObject.ListColumns("Status").Index).Value = "CALCULATED"
    End With
End Sub

Private Sub ValidateProductFamily(ByVal productId As String, ByVal family As String)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblProducts")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("ProductID").Index).Value), productId, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").Index).Value), family, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").Index).Value), "Yes", vbTextCompare) = 0 Then Exit Sub
    Next rowObject
    Err.Raise vbObjectError + 1305, "ValidateProductFamily", "Product and family do not match an active master row."
End Sub

Private Sub ValidateSize(ByVal family As String, ByVal basis As String, ByVal sizeMm As Double)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblSizes")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").Index).Value), family, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Basis").Index).Value), basis, vbTextCompare) = 0 _
            And Abs(CDbl(rowObject.Range.Cells(1, tableObject.ListColumns("SizeMM").Index).Value) - sizeMm) < 0.0001 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").Index).Value), "Yes", vbTextCompare) = 0 Then Exit Sub
    Next rowObject
    Err.Raise vbObjectError + 1306, "ValidateSize", "Size is not available for the selected family and basis."
End Sub

Private Sub ValidateThickness(ByVal family As String, ByVal thicknessMm As Double)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblThickness")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").Index).Value), family, vbTextCompare) = 0 _
            And Abs(CDbl(rowObject.Range.Cells(1, tableObject.ListColumns("ThicknessMM").Index).Value) - thicknessMm) < 0.0001 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").Index).Value), "Yes", vbTextCompare) = 0 Then Exit Sub
    Next rowObject
    Err.Raise vbObjectError + 1307, "ValidateThickness", "Thickness is not available for the selected family."
End Sub

Private Sub ValidateGrade(ByVal family As String, ByVal gradeId As String)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblGrades")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").Index).Value), family, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("GradeID").Index).Value), gradeId, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").Index).Value), "Yes", vbTextCompare) = 0 Then Exit Sub
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
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("PlantID").Index).Value), plantId, vbTextCompare) = 0 _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").Index).Value), "Yes", vbTextCompare) = 0 Then Exit Sub
    Next rowObject
    Err.Raise vbObjectError + 1310, "ValidatePlant", "Plant is not available in the active plant master."
End Sub
