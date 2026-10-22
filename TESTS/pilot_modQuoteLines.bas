Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modQuoteLines
' Purpose:
'   P0.4 Dynamic Quotation Line Items & Custom Sizing Engine
'   Supports adding, editing, clearing, and deleting quotation lines.
'   Supports coexistence of multiple distinct pipe sizes in a single quote.
'   Supports validated custom sizes with physical geometry checks.
'   Maintains contiguous numbering and recalculates line pricing.
' ============================================================

Public Sub AddQuoteLine(ByVal quoteId As String, ByVal productId As String, ByVal family As String, _
    ByVal basis As String, ByVal sizeMm As Double, ByVal thicknessMm As Double, _
    ByVal gradeId As String, ByVal quantity As Double, ByVal quantityUnit As String, _
    ByVal pipeLengthM As Double, ByVal plantId As String, _
    Optional ByVal legacyMode As Boolean = False, _
    Optional ByVal priceByMT As Boolean = False, _
    Optional ByVal manualFC As Variant, _
    Optional ByVal tripRate As Variant, _
    Optional ByVal isCustomSize As Boolean = False, _
    Optional ByVal extCoating As String = "", _
    Optional ByVal intCoating As String = "")

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
    
    Dim isCustom As Boolean
    isCustom = isCustomSize
    If Not isCustom Then
        On Error Resume Next
        ValidateSize family, basis, sizeMm
        If Err.Number <> 0 Then
            isCustom = True
            Err.Clear
        Else
            ValidateThickness family, thicknessMm
            If Err.Number <> 0 Then
                isCustom = True
                Err.Clear
            End If
        End If
        On Error GoTo 0
    End If
    
    If isCustom Then
        ValidateCustomDimensions family, basis, sizeMm, thicknessMm
    Else
        ValidateSize family, basis, sizeMm
        ValidateThickness family, thicknessMm
    End If
    
    ValidateGrade family, gradeId
    ValidateQuantityUnit quantityUnit
    ValidatePlant plantId

    Dim wsLines As Worksheet
    Set wsLines = tableObject.Parent
    modProtection.UnprotectSheet wsLines

    ' Maximum 10-Size Capacity Enforcement
    Dim currentActiveCount As Long
    Dim rCheck As ListRow
    currentActiveCount = 0
    For Each rCheck In tableObject.ListRows
        If Len(Trim$(CStr(rCheck.Range.Cells(1, tableObject.ListColumns("ProductID").index).value))) > 0 Then
            currentActiveCount = currentActiveCount + 1
        End If
    Next rCheck
    
    If currentActiveCount >= 10 Then
        Err.Raise vbObjectError + 1305, "AddQuoteLine", _
            "Quotation capacity reached: Maximum of 10 product-size lines are supported per quotation. (Currently " & CStr(currentActiveCount) & " of 10 sizes added). Please remove or edit an existing line."
    End If

    If tableObject.ListRows.count = 1 Then
        If Len(Trim$(CStr(tableObject.ListRows(1).Range.Cells(1, tableObject.ListColumns("ProductID").index).value))) = 0 Then
            Set lineObject = tableObject.ListRows(1)
        Else
            Set lineObject = tableObject.ListRows.Add
        End If
    Else
        Set lineObject = tableObject.ListRows.Add
    End If
    
    With lineObject.Range
        .Cells(1, tableObject.ListColumns("LineID").index).value = "L" & Format$(lineObject.index, "0000")
        .Cells(1, tableObject.ListColumns("QuoteID").index).value = quoteId
        .Cells(1, tableObject.ListColumns("ProductID").index).value = productId
        .Cells(1, tableObject.ListColumns("Family").index).value = family
        .Cells(1, tableObject.ListColumns("Basis").index).value = basis
        .Cells(1, tableObject.ListColumns("SizeMM").index).value = sizeMm
        .Cells(1, tableObject.ListColumns("ThicknessMM").index).value = thicknessMm
        .Cells(1, tableObject.ListColumns("GradeID").index).value = gradeId
        .Cells(1, tableObject.ListColumns("CoatingExternalID").index).value = extCoating
        .Cells(1, tableObject.ListColumns("CoatingInternalID").index).value = intCoating
        .Cells(1, tableObject.ListColumns("Quantity").index).value = quantity
        .Cells(1, tableObject.ListColumns("QuantityUnit").index).value = quantityUnit
        .Cells(1, tableObject.ListColumns("PipeLengthM").index).value = pipeLengthM
        .Cells(1, tableObject.ListColumns("PlantID").index).value = plantId
        .Cells(1, tableObject.ListColumns("Status").index).value = "DRAFT"
    End With

    On Error GoTo CalculationError

    If Len(extCoating) > 0 And UCase$(extCoating) <> "BARE" And UCase$(extCoating) <> "NONE" Then
        hasCoating = True
    ElseIf Len(intCoating) > 0 And UCase$(intCoating) <> "NONE" Then
        hasCoating = True
    Else
        hasCoating = (InStr(1, productId, "02", vbTextCompare) > 0 Or _
                      InStr(1, productId, "04", vbTextCompare) > 0)
    End If
    With lineData
        .basis = basis
        .sizeMm = sizeMm
        .wt = thicknessMm
        .qty = quantity
        .Unit = quantityUnit
        .PipeLen = pipeLengthM
        .PerTrailer = CDec(4)
        .hasExt = hasCoating
        .hasInt = hasCoating
    End With

    LoadDefaultRateInput rates, legacyMode
    rates.ExtRate = CDec(500)
    rates.IntRate = CDec(275)
    If IsMissing(tripRate) Then
        rates.tripRate = CDec(SettingValue("TripRateManual"))
    Else
        rates.tripRate = CDec(tripRate)
    End If
    If legacyMode Then
        If IsMissing(manualFC) Then
            rates.manualFC = CDec(SettingValue("DefaultFinancialPct"))
        Else
            rates.manualFC = CDec(manualFC)
        End If
    End If

    result = CalcLine(lineData, rates, priceByMT)
    WriteCalculatedQuoteLine lineObject, tableObject, result, priceByMT
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
    Exit Sub

CalculationError:
    On Error Resume Next
    lineObject.Delete
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
    On Error GoTo 0
    Err.Raise Err.Number, "AddQuoteLine", Err.Description
End Sub

Public Sub AddCustomQuoteLine(ByVal quoteId As String, ByVal productId As String, ByVal family As String, _
    ByVal basis As String, ByVal sizeMm As Double, ByVal thicknessMm As Double, _
    ByVal gradeId As String, ByVal quantity As Double, ByVal quantityUnit As String, _
    ByVal pipeLengthM As Double, ByVal plantId As String, _
    Optional ByVal extCoating As String = "Bare", Optional ByVal intCoating As String = "None")

    AddQuoteLine quoteId:=quoteId, productId:=productId, family:=family, _
                 basis:=basis, sizeMm:=sizeMm, thicknessMm:=thicknessMm, _
                 gradeId:=gradeId, quantity:=quantity, quantityUnit:=quantityUnit, _
                 pipeLengthM:=pipeLengthM, plantId:=plantId, isCustomSize:=True, _
                 extCoating:=extCoating, intCoating:=intCoating
End Sub

Public Sub DeleteQuoteLine(ByVal targetLineId As String)
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim rObj As ListRow
    Dim targetFound As Boolean
    Dim targetIdx As Long
    
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    modProtection.UnprotectSheet wsLines
    
    targetFound = False
    For Each rObj In tblLines.ListRows
        Dim lId As String
        lId = Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("LineID").index).value))
        If StrComp(lId, targetLineId, vbTextCompare) = 0 Then
            targetFound = True
            targetIdx = rObj.index
            Exit For
        End If
    Next rObj
    
    If targetFound Then
        If tblLines.ListRows.count = 1 Then
            ' If only 1 row, clear its contents instead of deleting table structure
            tblLines.ListRows(1).Range.ClearContents
        Else
            tblLines.ListRows(targetIdx).Delete
        End If
        RenumberLines
    Else
        Err.Raise vbObjectError + 1320, "DeleteQuoteLine", "Line ID " & targetLineId & " not found."
    End If
    
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
End Sub

Public Sub ClearQuoteLine(ByVal targetLineId As String)
    DeleteQuoteLine targetLineId
End Sub

Public Function DuplicateQuoteLine(ByVal sourceLineId As String) As String
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim rObj As ListRow
    Dim srcRow As ListRow
    
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    For Each rObj In tblLines.ListRows
        If StrComp(Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("LineID").index).value)), sourceLineId, vbTextCompare) = 0 Then
            Set srcRow = rObj
            Exit For
        End If
    Next rObj
    
    If srcRow Is Nothing Then
        Err.Raise vbObjectError + 1322, "DuplicateQuoteLine", "Source line ID " & sourceLineId & " not found."
    End If
    
    Dim qId As String, pId As String, fam As String, bas As String
    Dim sz As Double, wt As Double, grd As String, qty As Double, unt As String
    Dim pLen As Double, plt As String, leg As Boolean
    Dim extC As String, intC As String
    
    With srcRow.Range
        qId = Trim$(CStr(.Cells(1, tblLines.ListColumns("QuoteID").index).value))
        pId = Trim$(CStr(.Cells(1, tblLines.ListColumns("ProductID").index).value))
        fam = Trim$(CStr(.Cells(1, tblLines.ListColumns("Family").index).value))
        bas = Trim$(CStr(.Cells(1, tblLines.ListColumns("Basis").index).value))
        sz = Val(.Cells(1, tblLines.ListColumns("SizeMM").index).value)
        wt = Val(.Cells(1, tblLines.ListColumns("ThicknessMM").index).value)
        grd = Trim$(CStr(.Cells(1, tblLines.ListColumns("GradeID").index).value))
        extC = Trim$(CStr(.Cells(1, tblLines.ListColumns("CoatingExternalID").index).value))
        intC = Trim$(CStr(.Cells(1, tblLines.ListColumns("CoatingInternalID").index).value))
        qty = Val(.Cells(1, tblLines.ListColumns("Quantity").index).value)
        unt = Trim$(CStr(.Cells(1, tblLines.ListColumns("QuantityUnit").index).value))
        pLen = Val(.Cells(1, tblLines.ListColumns("PipeLengthM").index).value)
        plt = Trim$(CStr(.Cells(1, tblLines.ListColumns("PlantID").index).value))
    End With
    
    leg = (InStr(1, CStr(ThisWorkbook.Worksheets("QUOTE").Range("B17").value), "LEGACY", vbTextCompare) > 0 Or _
           InStr(1, CStr(ThisWorkbook.Worksheets("QUOTE").Range("E12").value), "LEGACY", vbTextCompare) > 0)
           
    AddQuoteLine quoteId:=qId, productId:=pId, family:=fam, basis:=bas, _
                 sizeMm:=sz, thicknessMm:=wt, gradeId:=grd, quantity:=qty, _
                 quantityUnit:=unt, pipeLengthM:=pLen, plantId:=plt, legacyMode:=leg, _
                 extCoating:=extC, intCoating:=intC
                 
    Dim newIdx As Long
    newIdx = tblLines.ListRows.count
    DuplicateQuoteLine = Trim$(CStr(tblLines.ListRows(newIdx).Range.Cells(1, tblLines.ListColumns("LineID").index).value))
End Function

Public Sub ClearAllQuoteLines()
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    modProtection.UnprotectSheet wsLines
    
    If tblLines.ListRows.count > 0 Then
        On Error Resume Next
        tblLines.DataBodyRange.Delete
        On Error GoTo 0
    End If
    
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
End Sub

Public Sub EditQuoteLine(ByVal targetLineId As String, _
                         Optional ByVal newQty As Variant, _
                         Optional ByVal newSizeMm As Variant, _
                         Optional ByVal newThkMm As Variant, _
                         Optional ByVal newPlantId As Variant)
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim rObj As ListRow
    Dim targetRow As ListRow
    
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    modProtection.UnprotectSheet wsLines
    
    For Each rObj In tblLines.ListRows
        If StrComp(Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("LineID").index).value)), targetLineId, vbTextCompare) = 0 Then
            Set targetRow = rObj
            Exit For
        End If
    Next rObj
    
    If targetRow Is Nothing Then
        Err.Raise vbObjectError + 1321, "EditQuoteLine", "Line ID " & targetLineId & " not found."
    End If
    
    With targetRow.Range
        If Not IsMissing(newQty) And Not IsEmpty(newQty) Then
            If Val(newQty) > 0 Then
                .Cells(1, tblLines.ListColumns("Quantity").index).value = Val(newQty)
            End If
        End If
        If Not IsMissing(newSizeMm) And Not IsEmpty(newSizeMm) Then
            If Val(newSizeMm) > 0 Then
                .Cells(1, tblLines.ListColumns("SizeMM").index).value = Val(newSizeMm)
            End If
        End If
        If Not IsMissing(newThkMm) And Not IsEmpty(newThkMm) Then
            If Val(newThkMm) > 0 Then
                .Cells(1, tblLines.ListColumns("ThicknessMM").index).value = Val(newThkMm)
            End If
        End If
        If Not IsMissing(newPlantId) And Not IsEmpty(newPlantId) Then
            If Len(Trim$(CStr(newPlantId))) > 0 Then
                .Cells(1, tblLines.ListColumns("PlantID").index).value = Trim$(CStr(newPlantId))
            End If
        End If
    End With
    
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
End Sub

Public Sub UpdateQuoteLineFull(ByVal targetLineId As String, ByVal productId As String, _
    ByVal family As String, ByVal basis As String, ByVal sizeMm As Double, _
    ByVal thicknessMm As Double, ByVal gradeId As String, ByVal quantity As Double, _
    ByVal quantityUnit As String, ByVal pipeLengthM As Double, ByVal plantId As String, _
    Optional ByVal extCoating As String = "", Optional ByVal intCoating As String = "", _
    Optional ByVal legacyMode As Boolean = False, Optional ByVal priceByMT As Boolean = False)

    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim rObj As ListRow
    Dim targetRow As ListRow
    Dim lineData As QuoteLineInput
    Dim rates As RateInput
    Dim result As LineResult
    Dim hasCoating As Boolean
    
    If Not IsValidQuantity(quantity) Then
        Err.Raise vbObjectError + 1302, "UpdateQuoteLineFull", "Quantity must be greater than zero."
    End If
    If Not IsValidPipeLength(pipeLengthM) Then
        Err.Raise vbObjectError + 1303, "UpdateQuoteLineFull", "Pipe length must be between 6 and 13 metres."
    End If
    If UCase$(Trim$(basis)) <> "OD" And UCase$(Trim$(basis)) <> "ID" Then
        Err.Raise vbObjectError + 1304, "UpdateQuoteLineFull", "Basis must be OD or ID."
    End If
    
    ValidateProductFamily productId, family
    
    Dim isCustom As Boolean
    isCustom = False
    On Error Resume Next
    ValidateSize family, basis, sizeMm
    If Err.Number <> 0 Then
        isCustom = True
        Err.Clear
    Else
        ValidateThickness family, thicknessMm
        If Err.Number <> 0 Then
            isCustom = True
            Err.Clear
        End If
    End If
    On Error GoTo 0
    
    If isCustom Then
        ValidateCustomDimensions family, basis, sizeMm, thicknessMm
    Else
        ValidateSize family, basis, sizeMm
        ValidateThickness family, thicknessMm
    End If
    
    ValidateGrade family, gradeId
    ValidateQuantityUnit quantityUnit
    ValidatePlant plantId

    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    modProtection.UnprotectSheet wsLines
    
    For Each rObj In tblLines.ListRows
        If StrComp(Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("LineID").index).value)), targetLineId, vbTextCompare) = 0 Then
            Set targetRow = rObj
            Exit For
        End If
    Next rObj
    
    If targetRow Is Nothing Then
        Err.Raise vbObjectError + 1321, "UpdateQuoteLineFull", "Line ID " & targetLineId & " not found."
    End If
    
    With targetRow.Range
        .Cells(1, tblLines.ListColumns("ProductID").index).value = productId
        .Cells(1, tblLines.ListColumns("Family").index).value = family
        .Cells(1, tblLines.ListColumns("Basis").index).value = basis
        .Cells(1, tblLines.ListColumns("SizeMM").index).value = sizeMm
        .Cells(1, tblLines.ListColumns("ThicknessMM").index).value = thicknessMm
        .Cells(1, tblLines.ListColumns("GradeID").index).value = gradeId
        .Cells(1, tblLines.ListColumns("CoatingExternalID").index).value = extCoating
        .Cells(1, tblLines.ListColumns("CoatingInternalID").index).value = intCoating
        .Cells(1, tblLines.ListColumns("Quantity").index).value = quantity
        .Cells(1, tblLines.ListColumns("QuantityUnit").index).value = quantityUnit
        .Cells(1, tblLines.ListColumns("PipeLengthM").index).value = pipeLengthM
        .Cells(1, tblLines.ListColumns("PlantID").index).value = plantId
    End With
    
    If Len(extCoating) > 0 And UCase$(extCoating) <> "BARE" And UCase$(extCoating) <> "NONE" Then
        hasCoating = True
    ElseIf Len(intCoating) > 0 And UCase$(intCoating) <> "NONE" Then
        hasCoating = True
    Else
        hasCoating = (InStr(1, productId, "02", vbTextCompare) > 0 Or InStr(1, productId, "04", vbTextCompare) > 0)
    End If
    
    With lineData
        .basis = basis
        .sizeMm = sizeMm
        .wt = thicknessMm
        .qty = quantity
        .Unit = quantityUnit
        .PipeLen = pipeLengthM
        .PerTrailer = CDec(4)
        .hasExt = hasCoating
        .hasInt = hasCoating
    End With
    
    LoadDefaultRateInput rates, legacyMode
    rates.ExtRate = CDec(500)
    rates.IntRate = CDec(275)
    rates.tripRate = CDec(SettingValue("TripRateManual"))
    If legacyMode Then
        rates.manualFC = CDec(SettingValue("DefaultFinancialPct"))
    End If
    
    result = CalcLine(lineData, rates, priceByMT)
    WriteCalculatedQuoteLine targetRow, tblLines, result, priceByMT
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
End Sub

Private Sub RenumberLines()
    Dim tblLines As ListObject
    Dim rObj As ListRow
    Dim idx As Long
    
    Set tblLines = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    idx = 1
    For Each rObj In tblLines.ListRows
        If Len(Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("ProductID").index).value))) > 0 Then
            rObj.Range.Cells(1, tblLines.ListColumns("LineID").index).value = "L" & Format$(idx, "0000")
            idx = idx + 1
        End If
    Next rObj
End Sub

Public Sub ValidateCustomDimensions(ByVal family As String, ByVal basis As String, ByVal sizeMm As Double, ByVal thicknessMm As Double)
    ' Physical sanity check for custom pipe manufacturing
    If sizeMm < 100# Or sizeMm > 3500# Then
        Err.Raise vbObjectError + 1311, "ValidateCustomDimensions", _
            "Custom pipe diameter must be between 100 mm and 3,500 mm (specified: " & Format$(sizeMm, "0.0") & " mm)."
    End If
    
    If thicknessMm < 3.5 Or thicknessMm > 60# Then
        Err.Raise vbObjectError + 1312, "ValidateCustomDimensions", _
            "Custom wall thickness must be between 3.5 mm and 60.0 mm (specified: " & Format$(thicknessMm, "0.0") & " mm)."
    End If
    
    If UCase$(Trim$(basis)) = "OD" Then
        If sizeMm <= (2# * thicknessMm) Then
            Err.Raise vbObjectError + 1313, "ValidateCustomDimensions", _
                "Invalid geometry: Outer Diameter (" & Format$(sizeMm, "0.0") & " mm) must be greater than twice wall thickness (" & Format$(2# * thicknessMm, "0.0") & " mm)."
        End If
    ElseIf UCase$(Trim$(basis)) = "ID" Then
        If sizeMm <= 0# Then
            Err.Raise vbObjectError + 1314, "ValidateCustomDimensions", _
                "Inner Diameter must be strictly positive."
        End If
    End If
End Sub

Private Sub WriteCalculatedQuoteLine(ByVal lineObject As ListRow, _
                                     ByVal tableObject As ListObject, _
                                     ByRef result As LineResult, _
                                     ByVal priceByMT As Boolean)
    With lineObject.Range
        .Cells(1, tableObject.ListColumns("CalculatedODMM").index).value = result.od
        .Cells(1, tableObject.ListColumns("CalculatedIDMM").index).value = result.ID
        .Cells(1, tableObject.ListColumns("MassKgPerM").index).value = result.MassPerM
        .Cells(1, tableObject.ListColumns("Metres").index).value = result.metres
        .Cells(1, tableObject.ListColumns("PipeCount").index).value = result.pipes
        .Cells(1, tableObject.ListColumns("WeightMT").index).value = result.mt
        .Cells(1, tableObject.ListColumns("UnitPrice").index).value = _
            IIf(priceByMT, result.TotalPerMT, result.TotalPerM)
        .Cells(1, tableObject.ListColumns("PriceUnit").index).value = _
            IIf(priceByMT, "MT", "Metres")
        .Cells(1, tableObject.ListColumns("LineTotal").index).value = result.LineValue
        .Cells(1, tableObject.ListColumns("Status").index).value = "CALCULATED"
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
            And (Abs(CDbl(rowObject.Range.Cells(1, tableObject.ListColumns("SizeMM").index).value) - sizeMm) <= 1#) _
            And StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Active").index).value), "Yes", vbTextCompare) = 0 Then Exit Sub
    Next rowObject
    Err.Raise vbObjectError + 1306, "ValidateSize", "Size is not available for the selected family and basis."
End Sub

Private Sub ValidateThickness(ByVal family As String, ByVal thicknessMm As Double)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Dim hasActiveCol As Boolean
    Dim actIdx As Long
    
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblThickness")
    On Error Resume Next
    actIdx = tableObject.ListColumns("Active").index
    hasActiveCol = (Err.Number = 0)
    On Error GoTo 0
    
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").index).value), family, vbTextCompare) = 0 _
            And (Abs(CDbl(rowObject.Range.Cells(1, tableObject.ListColumns("ThicknessMM").index).value) - thicknessMm) <= 0.05) Then
            If hasActiveCol Then
                If StrComp(CStr(rowObject.Range.Cells(1, actIdx).value), "Yes", vbTextCompare) = 0 Then Exit Sub
            Else
                Exit Sub
            End If
        End If
    Next rowObject
    Err.Raise vbObjectError + 1307, "ValidateThickness", "Thickness is not available for the selected family."
End Sub

Private Sub ValidateGrade(ByVal family As String, ByVal gradeId As String)
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("MASTERS").ListObjects("tblGrades")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("Family").index).value), family, vbTextCompare) = 0 _
            And (StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("GradeID").index).value), gradeId, vbTextCompare) = 0 _
                 Or StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("GradeName").index).value), gradeId, vbTextCompare) = 0 _
                 Or InStr(1, CStr(rowObject.Range.Cells(1, tableObject.ListColumns("GradeName").index).value), gradeId, vbTextCompare) > 0) _
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
    Dim hasActiveCol As Boolean
    Dim actIdx As Long
    
    Set tableObject = ThisWorkbook.Worksheets("PLANTS").ListObjects("tblPlants")
    On Error Resume Next
    actIdx = tableObject.ListColumns("Active").index
    hasActiveCol = (Err.Number = 0)
    On Error GoTo 0
    
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("PlantID").index).value), plantId, vbTextCompare) = 0 Then
            If hasActiveCol Then
                If StrComp(CStr(rowObject.Range.Cells(1, actIdx).value), "Yes", vbTextCompare) = 0 Then Exit Sub
            Else
                Exit Sub
            End If
        End If
    Next rowObject
    Err.Raise vbObjectError + 1310, "ValidatePlant", "Plant is not available in the active plant master."
End Sub

Public Function RunMultiSizeSelfTests() As String
    On Error GoTo FailureHandler
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim summary As QuoteSummaryData
    Dim header As QuoteHeaderData
    Dim testQuoteId As String
    Dim errCaught As Boolean
    
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    testQuoteId = "MKTG/ENQ/26-27/TEST-MULTI-10"
    
    modProtection.UnprotectSheet wsQ
    modProtection.UnprotectSheet wsLines
    
    ' Step 1: Start with empty table
    ClearAllQuoteLines
    
    ' Step 2: Test 1 line
    AddQuoteLine testQuoteId, "P02", "HSAW", "ID", 1226, 8.4, "HSAW-002", 1000, "Metres", 12, "P02"
    If tblLines.ListRows.count <> 1 Then
        RunMultiSizeSelfTests = "FAIL: Expected 1 quote line, got " & CStr(tblLines.ListRows.count)
        Exit Function
    End If
    header = modQuote.ReadQuoteSheet()
    summary = modQuote.CalcAllLines(header)
    If summary.ValidLineCount <> 1 Or CDbl(summary.TotalMetres) <> 1000# Then
        RunMultiSizeSelfTests = "FAIL: 1-line quotation calculation incorrect"
        Exit Function
    End If
    
    ' Step 3: Test 2 lines (different size and thickness)
    AddQuoteLine testQuoteId, "P02", "HSAW", "ID", 1300, 9.5, "HSAW-002", 500, "Metres", 12, "P02", isCustomSize:=True
    If tblLines.ListRows.count <> 2 Then
        RunMultiSizeSelfTests = "FAIL: Expected 2 quote lines, got " & CStr(tblLines.ListRows.count)
        Exit Function
    End If
    summary = modQuote.CalcAllLines(header)
    If summary.ValidLineCount <> 2 Or CDbl(summary.TotalMetres) <> 1500# Then
        RunMultiSizeSelfTests = "FAIL: 2-line quotation calculation incorrect"
        Exit Function
    End If
    
    ' Step 4: Add lines 3, 4, 5 (up to 5 distinct sizes)
    AddQuoteLine testQuoteId, "P02", "HSAW", "ID", 1350, 8.4, "HSAW-002", 600, "Metres", 12, "P02"
    AddQuoteLine testQuoteId, "P02", "HSAW", "ID", 1400, 8.4, "HSAW-002", 800, "Metres", 12, "P02"
    AddQuoteLine testQuoteId, "P02", "HSAW", "ID", 1500, 10#, "HSAW-002", 400, "Metres", 12, "P02", isCustomSize:=True
    
    If tblLines.ListRows.count <> 5 Then
        RunMultiSizeSelfTests = "FAIL: Expected 5 quote lines, got " & CStr(tblLines.ListRows.count)
        Exit Function
    End If
    summary = modQuote.CalcAllLines(header)
    If summary.ValidLineCount <> 5 Or CDbl(summary.TotalMetres) <> 3300# Then
        RunMultiSizeSelfTests = "FAIL: 5-line quotation total metres mismatch"
        Exit Function
    End If
    
    ' Step 5: Test Line Duplication (Duplicate Line 5 -> Line 6)
    Dim dupLineId As String
    dupLineId = DuplicateQuoteLine("L0005")
    If tblLines.ListRows.count <> 6 Then
        RunMultiSizeSelfTests = "FAIL: Duplication did not create line 6"
        Exit Function
    End If
    ' Verify independence: edit line 6 quantity without affecting line 5
    EditQuoteLine dupLineId, newQty:=200
    If CDbl(tblLines.ListRows(5).Range.Cells(1, tblLines.ListColumns("Quantity").index).value) <> 400# Then
        RunMultiSizeSelfTests = "FAIL: Line 5 quantity altered when editing duplicated line 6"
        Exit Function
    End If
    If CDbl(tblLines.ListRows(6).Range.Cells(1, tblLines.ListColumns("Quantity").index).value) <> 200# Then
        RunMultiSizeSelfTests = "FAIL: Line 6 quantity not updated to 200"
        Exit Function
    End If
    
    ' Step 6: Add lines 7, 8, 9, 10 (maximum 10 sizes capacity)
    AddQuoteLine testQuoteId, "P02", "HSAW", "ID", 1600, 10#, "HSAW-002", 300, "Metres", 12, "P02", isCustomSize:=True
    AddQuoteLine testQuoteId, "P02", "HSAW", "ID", 1700, 11#, "HSAW-002", 350, "Metres", 12, "P02", isCustomSize:=True
    AddQuoteLine testQuoteId, "P02", "HSAW", "ID", 1800, 12#, "HSAW-002", 250, "Metres", 12, "P02", isCustomSize:=True
    AddQuoteLine testQuoteId, "P02", "HSAW", "ID", 2000, 12#, "HSAW-002", 200, "Metres", 12, "P02", isCustomSize:=True
    
    If tblLines.ListRows.count <> 10 Then
        RunMultiSizeSelfTests = "FAIL: Expected exactly 10 quote lines, got " & CStr(tblLines.ListRows.count)
        Exit Function
    End If
    summary = modQuote.CalcAllLines(header)
    If summary.ValidLineCount <> 10 Then
        RunMultiSizeSelfTests = "FAIL: Expected 10 valid lines in 10-line quotation, got " & CStr(summary.ValidLineCount)
        Exit Function
    End If
    
    ' Step 7: Attempt 11th Line (MUST BE REJECTED)
    errCaught = False
    On Error Resume Next
    AddQuoteLine testQuoteId, "P02", "HSAW", "ID", 2200, 14#, "HSAW-002", 100, "Metres", 12, "P02", isCustomSize:=True
    If Err.Number <> 0 Then errCaught = True
    On Error GoTo FailureHandler
    
    If Not errCaught Then
        RunMultiSizeSelfTests = "FAIL: Did not reject 11th quotation line!"
        Exit Function
    End If
    
    ' Verify table remains intact with exactly 10 lines (no corruption)
    If tblLines.ListRows.count <> 10 Then
        RunMultiSizeSelfTests = "FAIL: 11th line attempt corrupted table line count: " & CStr(tblLines.ListRows.count)
        Exit Function
    End If
    
    ' Step 8: Remove a middle line (Line 5)
    DeleteQuoteLine "L0005"
    If tblLines.ListRows.count <> 9 Then
        RunMultiSizeSelfTests = "FAIL: Expected 9 lines after removing line 5, got " & CStr(tblLines.ListRows.count)
        Exit Function
    End If
    ' Verify lines are contiguous L0001 to L0009
    Dim chkIdx As Long
    For chkIdx = 1 To 9
        Dim expId As String, actId As String
        expId = "L" & Format$(chkIdx, "0000")
        actId = Trim$(CStr(tblLines.ListRows(chkIdx).Range.Cells(1, tblLines.ListColumns("LineID").index).value))
        If actId <> expId Then
            RunMultiSizeSelfTests = "FAIL: Non-contiguous line ID at index " & CStr(chkIdx) & ": expected " & expId & ", got " & actId
            Exit Function
        End If
    Next chkIdx
    
    ' Restore golden reference quotation
    modSelfTest.PopulateGoldenReferenceQuote
    RunMultiSizeSelfTests = "PASS"
    Exit Function

FailureHandler:
    Dim saveDesc As String
    saveDesc = Err.Description
    On Error Resume Next
    modSelfTest.PopulateGoldenReferenceQuote
    RunMultiSizeSelfTests = "FAIL: " & saveDesc
End Function

Public Function RunCustomSizeSelfTests() As String
    On Error GoTo FailureHandler
    Dim errCaught As Boolean
    
    ' 1. Test valid custom size
    ValidateCustomDimensions "HSAW", "ID", 1500, 10#
    ValidateCustomDimensions "HSAW", "OD", 1520, 10#
    
    ' 2. Test invalid diameter (< 100)
    errCaught = False
    On Error Resume Next
    ValidateCustomDimensions "HSAW", "OD", 50, 10#
    If Err.Number <> 0 Then errCaught = True
    On Error GoTo FailureHandler
    If Not errCaught Then
        RunCustomSizeSelfTests = "FAIL: Did not reject diameter < 100 mm"
        Exit Function
    End If
    
    ' 3. Test invalid thickness (< 3.5)
    errCaught = False
    On Error Resume Next
    ValidateCustomDimensions "HSAW", "OD", 1000, 2#
    If Err.Number <> 0 Then errCaught = True
    On Error GoTo FailureHandler
    If Not errCaught Then
        RunCustomSizeSelfTests = "FAIL: Did not reject thickness < 3.5 mm"
        Exit Function
    End If
    
    ' 4. Test physical geometry violation (OD <= 2*WT)
    errCaught = False
    On Error Resume Next
    ValidateCustomDimensions "HSAW", "OD", 100, 55#
    If Err.Number <> 0 Then errCaught = True
    On Error GoTo FailureHandler
    If Not errCaught Then
        RunCustomSizeSelfTests = "FAIL: Did not reject OD <= 2*WT"
        Exit Function
    End If
    
    RunCustomSizeSelfTests = "PASS"
    Exit Function

FailureHandler:
    RunCustomSizeSelfTests = "FAIL: " & Err.Description
End Function