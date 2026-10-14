Attribute VB_Name = "modQuote"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modQuote
' Purpose:
'   P6 Quote and pricing workflow
'   Wires QUOTE sheet and QUOTE_LINES to calculation engine
'   Updates live summary card on every change
' ============================================================

Public Type QuoteHeaderData
    QuoteID As String
    QuoteDate As Variant
    Customer As String
    ContactPerson As String
    DeliveryPlace As String
    PlantID As String
    PaymentTerms As String
    ValidityDays As Long
    PriceBasis As String
    LegacyMode As Boolean
End Type

Public Type QuoteSummaryData
    TotalMetres As Variant
    TotalWeightMT As Variant
    TotalPipeCount As Long
    TotalTrailers As Long
    TotalExclGST As Variant
    GSTAmount As Variant
    GrandTotalInclGST As Variant
    LineCount As Long
    ValidLineCount As Long
    Status As String
End Type

Public Function ReadQuoteSheet() As QuoteHeaderData
    Dim header As QuoteHeaderData
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("QUOTE")
    
    header.QuoteID = Trim$(CStr(ws.Range("B4").Value))
    header.QuoteDate = ws.Range("B5").Value
    header.Customer = Trim$(CStr(ws.Range("B8").Value))
    header.ContactPerson = Trim$(CStr(ws.Range("B9").Value))
    header.DeliveryPlace = Trim$(CStr(ws.Range("B10").Value))
    If Len(header.DeliveryPlace) = 0 Then header.DeliveryPlace = Trim$(CStr(ws.Range("B11").Value))
    header.PlantID = Trim$(CStr(ws.Range("B13").Value))
    If Len(header.PlantID) = 0 Then header.PlantID = Trim$(CStr(ws.Range("B14").Value))
    header.PaymentTerms = Trim$(CStr(ws.Range("B14").Value))
    If Len(header.PaymentTerms) = 0 Or IsNumeric(header.PaymentTerms) Then header.PaymentTerms = Trim$(CStr(ws.Range("B15").Value))
    header.ValidityDays = CLng(Val(ws.Range("B15").Value))
    If header.ValidityDays <= 0 Then header.ValidityDays = CLng(Val(ws.Range("B16").Value))
    header.PriceBasis = Trim$(CStr(ws.Range("B16").Value))
    If Len(header.PriceBasis) = 0 Or (header.PriceBasis <> "Metres" And header.PriceBasis <> "MT" And header.PriceBasis <> "Nos") Then
        header.PriceBasis = Trim$(CStr(ws.Range("B17").Value))
    End If
    If Len(header.PriceBasis) = 0 Then header.PriceBasis = "Metres"
    
    ' Check legacy mode switch from settings or header safely
    On Error Resume Next
    Dim vMode As Variant
    Dim tblSet As ListObject
    Dim rObj As ListRow
    vMode = Empty
    Set tblSet = ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")
    For Each rObj In tblSet.ListRows
        If StrComp(CStr(rObj.Range.Cells(1, tblSet.ListColumns("SettingKey").Index).Value), "LegacyMode", vbTextCompare) = 0 Then
            vMode = rObj.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value
            Exit For
        End If
    Next rObj
    header.LegacyMode = (UCase$(Trim$(CStr(vMode))) = "YES")
    On Error GoTo 0
    
    ReadQuoteSheet = header
End Function

Public Function CalcAllLines(ByRef header As QuoteHeaderData) As QuoteSummaryData
    Dim wsLines As Worksheet
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Dim summary As QuoteSummaryData
    Dim lineData As QuoteLineInput
    Dim rates As RateInput
    Dim result As LineResult
    Dim priceByMT As Boolean
    Dim hasExt As Boolean
    Dim hasInt As Boolean
    Dim prodId As String
    Dim basisVal As String
    Dim sizeVal As Double
    Dim wtVal As Double
    Dim qtyVal As Double
    Dim unitVal As String
    Dim pipeLenVal As Double
    Dim lineValid As Boolean
    Dim legacyMode As Boolean
    Dim candidateRows As Long
    
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tableObject = wsLines.ListObjects("tblQuoteLines")
    
    summary.TotalMetres = CDec(0)
    summary.TotalWeightMT = CDec(0)
    summary.TotalPipeCount = 0
    summary.TotalTrailers = 0
    summary.TotalExclGST = CDec(0)
    summary.GSTAmount = CDec(0)
    summary.GrandTotalInclGST = CDec(0)
    summary.LineCount = 0
    summary.ValidLineCount = 0
    summary.Status = "VALID"
    
    priceByMT = (UCase$(Trim$(header.PriceBasis)) = "MT")
    legacyMode = header.LegacyMode
    
    If tableObject.ListRows.Count = 0 Then
        summary.Status = "EMPTY"
        CalcAllLines = summary
        Exit Function
    End If
    
    candidateRows = 0
    For Each rowObject In tableObject.ListRows
        With rowObject.Range
            prodId = Trim$(CStr(.Cells(1, tableObject.ListColumns("ProductID").Index).Value))
            If Len(prodId) > 0 Then
                candidateRows = candidateRows + 1
            End If
        End With
    Next rowObject
    
    summary.LineCount = candidateRows
    If candidateRows = 0 Then
        summary.Status = "EMPTY"
        CalcAllLines = summary
        Exit Function
    End If
    
    For Each rowObject In tableObject.ListRows
        lineValid = True
        With rowObject.Range
            prodId = Trim$(CStr(.Cells(1, tableObject.ListColumns("ProductID").Index).Value))
            If Len(prodId) = 0 Then
                ' Skip blank placeholder row
            Else
                basisVal = Trim$(CStr(.Cells(1, tableObject.ListColumns("Basis").Index).Value))
                sizeVal = Val(.Cells(1, tableObject.ListColumns("SizeMM").Index).Value)
                wtVal = Val(.Cells(1, tableObject.ListColumns("ThicknessMM").Index).Value)
                qtyVal = Val(.Cells(1, tableObject.ListColumns("Quantity").Index).Value)
                unitVal = Trim$(CStr(.Cells(1, tableObject.ListColumns("QuantityUnit").Index).Value))
                pipeLenVal = Val(.Cells(1, tableObject.ListColumns("PipeLengthM").Index).Value)
                
                If (UCase$(basisVal) <> "OD" And UCase$(basisVal) <> "ID") _
                    Or sizeVal <= 0 Or wtVal <= 0 Or qtyVal <= 0 Then
                    lineValid = False
                End If
                If pipeLenVal < 6 Or pipeLenVal > 13 Then
                    pipeLenVal = 12#
                End If
                
                If lineValid Then
                    On Error GoTo LineCalcError
                    hasExt = (InStr(1, prodId, "02", vbTextCompare) > 0 Or InStr(1, prodId, "04", vbTextCompare) > 0)
                    hasInt = hasExt
                    
                    With lineData
                        .Basis = basisVal
                        .SizeMM = CDec(sizeVal)
                        .WT = CDec(wtVal)
                        .Qty = CDec(qtyVal)
                        .Unit = unitVal
                        .PipeLen = CDec(pipeLenVal)
                        .PerTrailer = CDec(4)
                        .HasExt = hasExt
                        .HasInt = hasInt
                    End With
                    
                    LoadDefaultRateInput rates, legacyMode
                    rates.ExtRate = CDec(500)
                    rates.IntRate = CDec(275)
                    If legacyMode Then
                        rates.TripRate = CDec(SettingValue("TripRateManual"))
                        rates.PI = CDec(3.14)
                        Dim manualFCs As Variant
                        manualFCs = Array(1542, 1534, 1529, 1525)
                        If rowObject.Index >= 1 And rowObject.Index <= 4 Then
                            rates.ManualFC = CDec(manualFCs(rowObject.Index - 1))
                        Else
                            rates.ManualFC = CDec(SettingValue("DefaultFinancialPct"))
                        End If
                    Else
                        rates.TripRate = CDec(125.32) * CDec(782)
                    End If
                    
                    result = CalcLine(lineData, rates, priceByMT)
                    
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
                    
                    summary.TotalMetres = summary.TotalMetres + CDec(result.Metres)
                    summary.TotalWeightMT = summary.TotalWeightMT + CDec(result.MT)
                    summary.TotalPipeCount = summary.TotalPipeCount + CLng(result.Pipes)
                    summary.TotalTrailers = summary.TotalTrailers + CLng(result.Trailers)
                    summary.TotalExclGST = summary.TotalExclGST + CDec(result.LineValue)
                    summary.ValidLineCount = summary.ValidLineCount + 1
                Else
                    .Cells(1, tableObject.ListColumns("Status").Index).Value = "INVALID"
                    summary.Status = "CONTAINS_INVALID_LINES"
                End If
            End If
        End With
    Next rowObject
    
    summary.GSTAmount = CDec(summary.TotalExclGST) * CDec("0.18")
    summary.GrandTotalInclGST = CDec(summary.TotalExclGST) + summary.GSTAmount
    CalcAllLines = summary
    Exit Function

LineCalcError:
    rowObject.Range.Cells(1, tableObject.ListColumns("Status").Index).Value = "ERROR: " & Err.Description
    summary.Status = "CALCULATION_ERROR"
    Resume Next
End Function

Public Sub UpdateLiveSummary(ByRef summary As QuoteSummaryData, ByRef header As QuoteHeaderData)
    Dim ws As Worksheet
    Dim prevAlerts As Boolean
    Set ws = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet ws
    
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    

    ' Title header
    If Not ws.Range("D3:E3").MergeCells Then
        ws.Range("D3:E3").Merge
    End If
    With ws.Range("D3")
        .Value = "LIVE QUOTATION SUMMARY"
        .Interior.Color = RGB(26, 82, 45) ' Deep Forest Green
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(3).RowHeight = 26
    
    ' Metrics
    ws.Range("D4").Value = "Total Metres"
    ws.Range("E4").Value = summary.TotalMetres
    ws.Range("E4").NumberFormat = "#,##0.00"
    
    ws.Range("D5").Value = "Total Weight (MT)"
    ws.Range("E5").Value = summary.TotalWeightMT
    ws.Range("E5").NumberFormat = "#,##0.000"
    
    ws.Range("D6").Value = "Total Pipe Count"
    ws.Range("E6").Value = summary.TotalPipeCount
    ws.Range("E6").NumberFormat = "#,##0"
    
    ws.Range("D7").Value = "Total Trailers"
    ws.Range("E7").Value = summary.TotalTrailers
    ws.Range("E7").NumberFormat = "#,##0"
    
    ws.Range("D8").Value = "Subtotal (Excl. GST)"
    ws.Range("E8").Value = summary.TotalExclGST
    ws.Range("E8").NumberFormat = "#,##0.00"
    
    ws.Range("D9").Value = "GST @ 18%"
    ws.Range("E9").Value = summary.GSTAmount
    ws.Range("E9").NumberFormat = "#,##0.00"
    
    ws.Range("D10").Value = "Grand Total (Incl. GST)"
    ws.Range("E10").Value = summary.GrandTotalInclGST
    ws.Range("E10").NumberFormat = "#,##0.00"
    
    ws.Range("D11").Value = "Workflow Status"
    Dim curWf As String
    curWf = Trim$(CStr(ws.Range("E11").Value))
    If summary.LineCount = 0 Or summary.ValidLineCount = 0 Then
        ws.Range("E11").Value = "Draft"
    ElseIf curWf = "Approved" Or curWf = "Issued" Then
        modApproval.InvalidateApprovalIfModified "Recalculation detected quote edits"
    ElseIf curWf = "" Or curWf = "Draft" Or curWf = "VALID" Or curWf = "EMPTY" Then
        ws.Range("E11").Value = "Calculated"
    End If
    
    ws.Range("D12").Value = "Mode"
    ws.Range("E12").Value = IIf(header.LegacyMode, "LEGACY", "NORMAL (Option 1)")
    
    ws.Range("D13").Value = "Supplying Plant"
    ws.Range("E13").Value = header.PlantID
    
    ws.Range("D14").Value = "Valid Lines"
    ws.Range("E14").NumberFormat = "@"
    ws.Range("E14").Value = CStr(summary.ValidLineCount) & " of " & CStr(summary.LineCount)
    
    ' Professional formatting with elegant contrast and generous spacing
    Dim r As Long
    For r = 4 To 14
        ws.Rows(r).RowHeight = 21
        If r = 10 Then
            ' Grand Total Row: High-visibility soft gold card accent with rich amber text
            ws.Range("D10:E10").Interior.Color = RGB(255, 243, 205) ' Soft warm gold
            ws.Range("D10:E10").Font.Color = RGB(120, 60, 0)       ' Rich dark amber
            ws.Range("D10:E10").Font.Bold = True
            ws.Range("D10:E10").Font.Size = 11
        ElseIf r = 8 Then
            ' Subtotal row
            ws.Range("D8:E8").Interior.Color = RGB(242, 248, 242)   ' Soft mint
            ws.Range("D8:E8").Font.Color = RGB(20, 50, 20)
            ws.Range("D8:E8").Font.Bold = True
            ws.Range("D8:E8").Font.Size = 10
        Else
            ws.Range("D" & r & ":E" & r).Interior.Color = IIf(r Mod 2 = 0, RGB(250, 252, 250), RGB(255, 255, 255))
            ws.Range("D" & r & ":E" & r).Font.Color = RGB(40, 40, 40)
            ws.Range("D" & r & ":E" & r).Font.Bold = False
            ws.Range("D" & r & ":E" & r).Font.Size = 10
        End If
        ws.Range("D" & r).HorizontalAlignment = xlLeft
        ws.Range("E" & r).HorizontalAlignment = xlRight
    Next r
    ws.Range("E11").HorizontalAlignment = xlCenter
    ws.Range("E12").HorizontalAlignment = xlCenter
    ws.Range("E14").HorizontalAlignment = xlCenter
    
    ' Subtle refined borders
    With ws.Range("D3:E14").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(190, 215, 190)
    End With
    
    ' Leave cells unlocked and editable per user request
    ws.Range("D3:E14").Locked = False
    
    Application.DisplayAlerts = prevAlerts
End Sub

Public Sub RecalculateQuote()
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousDisplayAlerts As Boolean
    Dim header As QuoteHeaderData
    Dim summary As QuoteSummaryData
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    
    On Error GoTo ErrorHandler
    
    previousEnableEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousDisplayAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    modProtection.UnprotectSheet wsQ
    modProtection.UnprotectSheet wsLines
    
    header = ReadQuoteSheet()
    summary = CalcAllLines(header)
    UpdateLiveSummary summary, header
    
CleanExit:
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    Application.DisplayAlerts = previousDisplayAlerts
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub

ErrorHandler:
    On Error Resume Next
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    On Error GoTo 0
    Application.DisplayAlerts = previousDisplayAlerts
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    Err.Raise Err.Number, "RecalculateQuote", Err.Description
End Sub

Public Sub ClearQuoteLines()
    Dim wsLines As Worksheet
    Dim wsQ As Worksheet
    Dim tableObject As ListObject
    Dim prevEvents As Boolean
    Dim prevAlerts As Boolean
    Dim r As Long
    
    On Error GoTo ErrorHandler
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet wsLines
    modProtection.UnprotectSheet wsQ
    
    Set tableObject = wsLines.ListObjects("tblQuoteLines")
    If Not tableObject.DataBodyRange Is Nothing Then
        tableObject.DataBodyRange.ClearContents
        For r = tableObject.ListRows.Count To 2 Step -1
            tableObject.ListRows(r).Delete
        Next r
    End If
    
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = prevEvents
    RecalculateQuote
    Exit Sub
ErrorHandler:
    On Error Resume Next
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    On Error GoTo 0
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = prevEvents
    Err.Raise Err.Number, "ClearQuoteLines", Err.Description
End Sub

Public Sub DeleteQuoteLine(ByVal lineId As String)
    Dim wsLines As Worksheet
    Dim wsQ As Worksheet
    Dim tableObject As ListObject
    Dim rowObject As ListRow
    Dim prevEvents As Boolean
    Dim prevAlerts As Boolean
    
    On Error GoTo ErrorHandler
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    modProtection.UnprotectSheet wsLines
    modProtection.UnprotectSheet wsQ
    
    Set tableObject = wsLines.ListObjects("tblQuoteLines")
    For Each rowObject In tableObject.ListRows
        If StrComp(CStr(rowObject.Range.Cells(1, tableObject.ListColumns("LineID").Index).Value), lineId, vbTextCompare) = 0 Then
            If tableObject.ListRows.Count = 1 Then
                tableObject.DataBodyRange.ClearContents
            Else
                rowObject.Delete
            End If
            Exit For
        End If
    Next rowObject
    
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = prevEvents
    RecalculateQuote
    Exit Sub
ErrorHandler:
    On Error Resume Next
    modProtection.ProtectSheet wsLines, userInterfaceOnly:=True
    modProtection.ProtectSheet wsQ, userInterfaceOnly:=True
    On Error GoTo 0
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = prevEvents
    Err.Raise Err.Number, "DeleteQuoteLine", Err.Description
End Sub

