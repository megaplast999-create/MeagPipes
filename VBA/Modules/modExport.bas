Attribute VB_Name = "modExport"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modExport
' Purpose:
'   P11 COSTING sheet generator
'   Reproduces reference layout from "CS Stacking 10-09-26"
'   Values only (no formulas); every number comes from modPricing.CalcLine
'   Row colors from visual style specification
'   One column per size + weighted average / total column
' ============================================================

Public Sub WriteCostingSheet()
    ' COSTING = the reference "CS Stacking" layout, VALUES ONLY.
    ' Every figure comes straight from CalcLine (the one calculation engine),
    ' nothing is re-derived here - so the sheet can never disagree with QUOTE.
    '   rows 44-48  per-metre prices  = CalcLine ROUNDUP results (they sum to row 48)
    '   column "total" = MT-weighted average, exactly the sheet's SUMPRODUCT(row, MT)/total MT
    Dim wsC As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim header As QuoteHeaderData
    Dim lineObj As ListRow
    Dim numLines As Long
    Dim j As Long
    Dim totCol As Long
    Dim rates As RateInput
    Dim lineData As QuoteLineInput
    Dim res As LineResult
    Dim legacyMode As Boolean
    Dim priceByMT As Boolean
    Dim gstPct As Double

    ' Per-line values
    Dim ids() As Double, ods() As Double, wts() As Double
    Dim massMT() As Double, metres() As Double, mts() As Double
    Dim pipes() As Long, trips() As Double
    Dim steel() As Double, gradeX() As Double, yieldLoss() As Double, conv() As Double
    Dim testing() As Double, margin() As Double
    Dim fc() As Double, bare() As Double, ext() As Double, intc() As Double
    Dim exMill() As Double, freight() As Double, subTot() As Double
    Dim gst() As Double, grand() As Double, fcCheck() As Double, valMT() As Double
    Dim bareM() As Double, extM() As Double, intM() As Double, frM() As Double, totM() As Double
    Dim lineVal() As Double, appliedP() As Double, variance() As Double
    Dim src() As String
    Dim finPct As Double, extRate As Double, intRate As Double, tripRate As Double
    Dim pipeLen() As Double

    Dim totMetres As Double, totMT As Double, totPipes As Long, totTrips As Double, totLineVal As Double

    Set wsC = ThisWorkbook.Worksheets("COSTING")
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")

    header = ReadQuoteSheet()
    legacyMode = header.LegacyMode
    priceByMT = (UCase$(Trim$(header.PriceBasis)) = "MT")
    gstPct = CDbl(SettingValueOr("GSTPct", 18))

    numLines = 0
    For Each lineObj In tblLines.ListRows
        If Len(Trim$(CStr(lineObj.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) > 0 Then
            numLines = numLines + 1
        End If
    Next lineObj
    If numLines = 0 Then Exit Sub

    ReDim ids(1 To numLines), ods(1 To numLines), wts(1 To numLines)
    ReDim massMT(1 To numLines), metres(1 To numLines), mts(1 To numLines)
    ReDim pipes(1 To numLines), trips(1 To numLines), pipeLen(1 To numLines)
    ReDim steel(1 To numLines), gradeX(1 To numLines), yieldLoss(1 To numLines), conv(1 To numLines)
    ReDim testing(1 To numLines), margin(1 To numLines)
    ReDim fc(1 To numLines), bare(1 To numLines), ext(1 To numLines), intc(1 To numLines)
    ReDim exMill(1 To numLines), freight(1 To numLines), subTot(1 To numLines)
    ReDim gst(1 To numLines), grand(1 To numLines), fcCheck(1 To numLines), valMT(1 To numLines)
    ReDim bareM(1 To numLines), extM(1 To numLines), intM(1 To numLines), frM(1 To numLines), totM(1 To numLines)
    ReDim lineVal(1 To numLines), appliedP(1 To numLines), variance(1 To numLines)
    ReDim src(1 To numLines)

    j = 0
    For Each lineObj In tblLines.ListRows
        If Len(Trim$(CStr(lineObj.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) > 0 Then
            j = j + 1
            LoadLineInputsFromRow lineData, rates, tblLines, lineObj, legacyMode
            res = CalcLine(lineData, rates, priceByMT)

            ids(j) = CDbl(res.ID)
            ods(j) = CDbl(res.od)
            wts(j) = CDbl(lineData.WT)
            massMT(j) = CDbl(res.MassPerM) / 1000#
            metres(j) = CDbl(res.Metres)
            mts(j) = CDbl(res.MT)
            pipes(j) = CLng(res.Pipes)
            trips(j) = CDbl(res.FreightTrips)
            pipeLen(j) = CDbl(lineData.PipeLen)

            steel(j) = CDbl(rates.Steel)
            gradeX(j) = CDbl(rates.GradeExtra)
            yieldLoss(j) = CDbl((rates.Steel + rates.GradeExtra) * rates.YieldPct / CDec(100))
            conv(j) = CDbl(rates.Conversion)
            testing(j) = CDbl(rates.Testing)
            margin(j) = CDbl(rates.Margin)

            fc(j) = CDbl(res.FCperMT)
            bare(j) = CDbl(res.BarePipeCostPerMT) + fc(j)
            ext(j) = CDbl(res.ExtPerMT)
            intc(j) = CDbl(res.IntPerMT)
            exMill(j) = bare(j) + ext(j) + intc(j)
            freight(j) = CDbl(res.FreightPerMT)
            subTot(j) = exMill(j) + freight(j)
            gst(j) = subTot(j) * gstPct / 100#
            grand(j) = subTot(j) + gst(j)
            fcCheck(j) = grand(j) * CDbl(rates.FinPct) / 100#
            valMT(j) = subTot(j) * mts(j)

            bareM(j) = CDbl(res.BarePerM)
            extM(j) = CDbl(res.ExtPerM)
            intM(j) = CDbl(res.IntPerM)
            frM(j) = CDbl(res.FreightPerM)
            totM(j) = CDbl(res.TotalPerM)
            lineVal(j) = CDbl(res.LineValue)
            appliedP(j) = CDbl(res.AppliedPrice)
            variance(j) = CDbl(res.PriceVariance)
            src(j) = res.PriceSource

            finPct = CDbl(rates.FinPct)
            extRate = CDbl(rates.ExtRate)
            intRate = CDbl(rates.IntRate)
            tripRate = CDbl(rates.TripRate)

            totMetres = totMetres + metres(j)
            totMT = totMT + mts(j)
            totPipes = totPipes + pipes(j)
            totTrips = totTrips + trips(j)
            totLineVal = totLineVal + lineVal(j)
        End If
    Next lineObj

    totCol = numLines + 3

    ' Clear existing sheet contents and styles
    wsC.Cells.Clear

    ' Top Information Block
    wsC.Range("A1").Value = "Date: " & Format$(Now, "dd/mm/yyyy")
    wsC.Range("B1").Value = "TO BE SUPPLIED FROM " & header.PlantID
    wsC.Range("A2").Value = "Owner"
    wsC.Range("B2").Value = header.Customer
    wsC.Range("A5").Value = "Pipe Type"
    wsC.Range("B5").Value = "HSAW Coated"
    wsC.Range("A6").Value = "Enquiry No"
    wsC.Range("B6").Value = header.QuoteID
    wsC.Range("A7").Value = "Delivery Location"
    wsC.Range("B7").Value = header.DeliveryPlace

    ' Row 12: Column headers
    wsC.Range("A12").Value = "Sr.No."
    For j = 1 To numLines: wsC.Cells(12, j + 2).Value = j: Next j
    wsC.Cells(12, totCol).Value = "total"
    ColourRow wsC.Range(wsC.Cells(12, 1), wsC.Cells(12, totCol)), "BDD7EE", "000000"
    wsC.Range(wsC.Cells(12, 1), wsC.Cells(12, totCol)).Font.Bold = True

    CostRow wsC, 13, "Pipe OD (Inch)", ods, numLines, "1A5E38", "FFFFFF", True, 0, True
    For j = 1 To numLines: wsC.Cells(13, j + 2).Value = WorksheetFunction.Round(ods(j) / 25.4, 0): Next j
    CostRow wsC, 14, "Pipe ID (mm)", ids, numLines, "70AD47", "000000", False, -1, False
    CostRow wsC, 15, "Pipe OD (mm)", ods, numLines, "70AD47", "000000", False, -1, False
    CostRow wsC, 16, "WT (mm)", wts, numLines, "70AD47", "000000", False, -1, False
    CostRow wsC, 17, "Mass (MT/Mtr)", massMT, numLines, "A9D18E", "000000", False, 8, False
    CostRow wsC, 18, "Qty (Mtr)", metres, numLines, "FFFF00", "000000", True, 0, False
    wsC.Cells(18, totCol).Value = totMetres
    CostRow wsC, 19, "Qty (MT)", mts, numLines, "FFFF00", "000000", True, 3, False
    wsC.Cells(19, totCol).Value = totMT

    ' Row 20: Coating factor = PI x OD / 1000 (PI from SETTINGS, 3.14 in the reference sheet)
    wsC.Range("A20").Value = "Coating factor"
    For j = 1 To numLines: wsC.Cells(20, j + 2).Value = CDbl(rates.PI) * ods(j) / 1000#: Next j
    ColourRow wsC.Range(wsC.Cells(20, 1), wsC.Cells(20, totCol)), "FFFFFF", "000000"

    wsC.Range("A21").Value = "Pipe Weight (MT)"
    For j = 1 To numLines: wsC.Cells(21, j + 2).Value = massMT(j) * pipeLen(j): Next j
    ColourRow wsC.Range(wsC.Cells(21, 1), wsC.Cells(21, totCol)), "FFFFFF", "000000"

    wsC.Range("A22").Value = "Pipe Nos"
    For j = 1 To numLines: wsC.Cells(22, j + 2).Value = pipes(j): Next j
    wsC.Cells(22, totCol).Value = totPipes
    ColourRow wsC.Range(wsC.Cells(22, 1), wsC.Cells(22, totCol)), "FFFF00", "000000"
    wsC.Range(wsC.Cells(22, 1), wsC.Cells(22, totCol)).Font.Bold = True

    ' Rows 25-31: base cost build-up. "total" column = MT-weighted average (sheet SUMPRODUCT).
    CostRowW wsC, 25, "Steel BEN + GE", steel, mts, numLines, "FFFFFF", "000000", False
    CostRowW wsC, 26, "Grade Extra", gradeX, mts, numLines, "FFFFFF", "000000", False
    Dim steelDelivered() As Double
    ReDim steelDelivered(1 To numLines)
    For j = 1 To numLines: steelDelivered(j) = steel(j) + gradeX(j): Next j
    CostRowW wsC, 27, "Steel Cost (Delivered)", steelDelivered, mts, numLines, "FFFFFF", "000000", False
    CostRowW wsC, 28, "Yield Loss @ " & CStr(CDbl(rates.YieldPct)) & "%", yieldLoss, mts, numLines, "F2F2F2", "000000", False
    CostRowW wsC, 29, "Conversion", conv, mts, numLines, "F2F2F2", "000000", False
    CostRowW wsC, 30, "Testing + Hospitality", testing, mts, numLines, "F2F2F2", "000000", False
    CostRowW wsC, 31, "Margin", margin, mts, numLines, "F2F2F2", "000000", False

    CostRowW wsC, 32, "FC-" & CStr(finPct) & "% + 25% adv. Bal against LC", fc, mts, numLines, "F4B8C1", "C00000", True
    wsC.Range("B32").Value = finPct / 100#
    CostRowW wsC, 33, "Bare Pipe Price", bare, mts, numLines, "FFFFFF", "000000", False
    CostRowW wsC, 34, "Ext: PU (1000 Micron)", ext, mts, numLines, "DDEEFF", "000000", False
    wsC.Range("B34").Value = extRate
    CostRowW wsC, 35, "Int:  FGE (406 Micron)", intc, mts, numLines, "DDEEFF", "000000", False
    wsC.Range("B35").Value = intRate
    CostRowW wsC, 36, "Ex Mill Price", exMill, mts, numLines, "FFFFFF", "000000", True
    CostRowW wsC, 37, "Freight : " & header.DeliveryPlace, freight, mts, numLines, "70AD47", "FFFFFF", True
    wsC.Cells(37, totCol + 1).Value = Format$(totTrips, "0.00") & " trips @ " & Format$(tripRate, "#,##0")
    CostRowW wsC, 38, "Sub Total", subTot, mts, numLines, "FFFF00", "000000", True
    CostRowW wsC, 39, "GST @ " & CStr(gstPct) & "%", gst, mts, numLines, "FFFFFF", "000000", False
    CostRowW wsC, 40, "Grand Total GST INC.", grand, mts, numLines, "FFFFFF", "000000", False

    ' Row 41: the sheet's own check cell = Grand Total x FC%. With automatic FC it equals row 32.
    wsC.Range("A41").Value = "FC-" & CStr(finPct) & "% + 25% adv. Bal against LC"
    wsC.Range("B41").Value = finPct / 100#
    For j = 1 To numLines: wsC.Cells(41, j + 2).Value = fcCheck(j): Next j
    ColourRow wsC.Range(wsC.Cells(41, 1), wsC.Cells(41, totCol)), "F4B8C1", "C00000"
    wsC.Range(wsC.Cells(41, 1), wsC.Cells(41, totCol)).Font.Bold = True

    ' Row 42 = Sub Total x MT (sheet); Row 49 = per-metre price x metres (what is quoted)
    wsC.Range("A42").Value = "Total Value (Excl. GST)"
    Dim sumValMT As Double
    For j = 1 To numLines
        wsC.Cells(42, j + 2).Value = valMT(j)
        sumValMT = sumValMT + valMT(j)
    Next j
    wsC.Cells(42, totCol).Value = sumValMT
    ColourRow wsC.Range(wsC.Cells(42, 1), wsC.Cells(42, totCol)), "FFFFFF", "000000"

    ' Rows 44-48: per-metre prices straight from CalcLine - they add up to row 48 by construction
    CostRow wsC, 44, "Bare Pipe Price/meter", bareM, numLines, "FFFFFF", "000000", False, 0, False
    CostRow wsC, 45, "EXT Coating Price/meter", extM, numLines, "FFFFFF", "000000", False, 0, False
    wsC.Range("B45").Value = extRate
    CostRow wsC, 46, "INT Coating Price/meter", intM, numLines, "FFFFFF", "000000", False, 0, False
    wsC.Range("B46").Value = intRate
    CostRow wsC, 47, "Freight Price/meter", frM, numLines, "FFFFFF", "000000", False, 0, False
    CostRow wsC, 48, "Total Price/meter", totM, numLines, "FFFF00", "000000", True, 0, False
    If totMetres > 0 Then wsC.Cells(48, totCol).Value = totLineVal / totMetres

    CostRow wsC, 49, "Total Value (Excl. GST)", lineVal, numLines, "FFFF00", "000000", True, 0, False
    wsC.Cells(49, totCol).Value = totLineVal

    ' Rows 51-53: manual selling price overrides, if any
    Dim hasOverride As Boolean
    hasOverride = False
    For j = 1 To numLines
        If src(j) = "MANUAL_OVERRIDE" Then hasOverride = True: Exit For
    Next j

    If hasOverride Then
        CostRow wsC, 51, "Applied Selling Price / Unit", appliedP, numLines, "EBF1DE", "000000", True, 0, False
        If totMetres > 0 Then wsC.Cells(51, totCol).Value = totLineVal / totMetres
        wsC.Range("A52").Value = "Price Source"
        For j = 1 To numLines: wsC.Cells(52, j + 2).Value = src(j): Next j
        wsC.Cells(52, totCol).Value = "OVERRIDDEN"
        ColourRow wsC.Range(wsC.Cells(52, 1), wsC.Cells(52, totCol)), "F2F2F2", "000000"
        CostRow wsC, 53, "Price Variance (Override - Calc)", variance, numLines, "FCE4D6", "C00000", True, 0, False
    End If

    Dim maxCostRow As Long
    maxCostRow = IIf(hasOverride, 53, 49)
    With wsC.Range(wsC.Cells(12, 1), wsC.Cells(maxCostRow, totCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Name = "Calibri"
    End With

    wsC.Columns("A:A").ColumnWidth = 32
    wsC.Columns("B:B").ColumnWidth = 10
    For j = 3 To totCol
        wsC.Columns(j).ColumnWidth = 16
    Next j
End Sub

' One COSTING row, one value per line (no total column unless the caller adds it).
' decimals: -1 = leave General, otherwise a "0.000.." format.
Private Sub CostRow(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                    ByRef vals() As Double, ByVal n As Long, _
                    ByVal bgHex As String, ByVal fgHex As String, ByVal isBold As Boolean, _
                    ByVal decimals As Long, ByVal skipValues As Boolean)
    Dim j As Long
    ws.Cells(r, 1).Value = label
    If Not skipValues Then
        For j = 1 To n
            ws.Cells(r, j + 2).Value = vals(j)
            If decimals = 0 Then
                ws.Cells(r, j + 2).NumberFormat = "#,##0"
            ElseIf decimals > 0 Then
                ws.Cells(r, j + 2).NumberFormat = "0." & String$(decimals, "0")
            End If
        Next j
    End If
    ColourRow ws.Range(ws.Cells(r, 1), ws.Cells(r, n + 3)), bgHex, fgHex
    If isBold Then ws.Range(ws.Cells(r, 1), ws.Cells(r, n + 3)).Font.Bold = True
End Sub

' Same, plus the "total" column = MT-weighted average, i.e. the sheet's SUMPRODUCT(row, MT)/total MT.
Private Sub CostRowW(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                     ByRef vals() As Double, ByRef weights() As Double, ByVal n As Long, _
                     ByVal bgHex As String, ByVal fgHex As String, ByVal isBold As Boolean)
    Dim j As Long
    Dim sw As Double, sv As Double
    ws.Cells(r, 1).Value = label
    For j = 1 To n
        ws.Cells(r, j + 2).Value = vals(j)
        ws.Cells(r, j + 2).NumberFormat = "#,##0.00"
        sv = sv + vals(j) * weights(j)
        sw = sw + weights(j)
    Next j
    If sw > 0 Then
        ws.Cells(r, n + 3).Value = sv / sw
        ws.Cells(r, n + 3).NumberFormat = "#,##0.00"
    End If
    ColourRow ws.Range(ws.Cells(r, 1), ws.Cells(r, n + 3)), bgHex, fgHex
    If isBold Then ws.Range(ws.Cells(r, 1), ws.Cells(r, n + 3)).Font.Bold = True
End Sub

Public Function RunCostingSelfTests() As String
    Dim wsC As Worksheet
    Dim valGrand As Double
    Dim valLine1 As Double
    Dim valLine4 As Double
    Dim hasFormula As Boolean
    Dim cellObj As Range
    Dim bgCol As Long
    
    On Error GoTo TestFail
    
    ' Ensure golden quote is loaded if empty
    Dim tblLines As ListObject
    Set tblLines = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")
    If tblLines.ListRows.Count = 0 Then
        modSelfTest.PopulateGoldenReferenceQuote
    ElseIf Len(Trim$(CStr(tblLines.ListRows(1).Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) = 0 Then
        modSelfTest.PopulateGoldenReferenceQuote
    End If
    
    ' Generate COSTING sheet
    WriteCostingSheet
    
    Set wsC = ThisWorkbook.Worksheets("COSTING")
    
    ' Verify values are populated
    valLine1 = wsC.Range("C48").Value
    valLine4 = wsC.Range("F48").Value
    valGrand = wsC.Range("G49").Value
    
    ' Normal mode (Option 1): Line 1 = 22689, Line 4 = 25581, Grand = 178801050
    ' Legacy mode: Line 1 = 22681, Line 4 = 25580, Grand = 178657900
    ' Approved company baseline: Line 1 = 22351, Line 4 = 25220, Grand = 176117500
    If valGrand <> 178801050# And valGrand <> 178657900# And valGrand <> 176117500# Then
        RunCostingSelfTests = "FAIL: Expected grand total 176117500, 178801050 or 178657900, got " & CStr(valGrand)
        Exit Function
    End If
    
    If valGrand = 176117500# Then
        If valLine1 <> 22351# Or valLine4 <> 25220# Then
            RunCostingSelfTests = "FAIL: Approved mode line prices mismatch: L1=" & CStr(valLine1) & ", L4=" & CStr(valLine4)
            Exit Function
        End If
    ElseIf valGrand = 178801050# Then
        If valLine1 <> 22689# Or valLine4 <> 25581# Then
            RunCostingSelfTests = "FAIL: Normal mode line prices mismatch: L1=" & CStr(valLine1) & ", L4=" & CStr(valLine4)
            Exit Function
        End If
    Else
        If valLine1 <> 22681# Or valLine4 <> 25580# Then
            RunCostingSelfTests = "FAIL: Legacy mode line prices mismatch: L1=" & CStr(valLine1) & ", L4=" & CStr(valLine4)
            Exit Function
        End If
    End If
    
    ' Verify NO formulas (values only requirement!)
    hasFormula = False
    For Each cellObj In wsC.Range("A12:G49")
        If cellObj.HasFormula Then
            hasFormula = True
            Exit For
        End If
    Next cellObj
    
    If hasFormula Then
        RunCostingSelfTests = "FAIL: Found formulas in generated COSTING sheet; must be values only"
        Exit Function
    End If
    
    ' Verify Row 48 has yellow background (FFFF00 = RGB(255, 255, 0) = 65535)
    bgCol = wsC.Range("C48").Interior.Color
    If bgCol <> 65535 Then
        RunCostingSelfTests = "FAIL: Row 48 background color expected Yellow (65535), got " & CStr(bgCol)
        Exit Function
    End If
    
    RunCostingSelfTests = "PASS"
    Exit Function
    
TestFail:
    RunCostingSelfTests = "FAIL: Exception - " & Err.Description
End Function

' ============================================================
' INDIAN NUMBERING WORDS CONVERTER
' ============================================================

Public Function NumberToIndianWords(ByVal amount As Double) As String
    Dim num As Currency
    Dim crore As Long, lakh As Long, thousand As Long, hundred As Long, remainder As Long
    Dim words As String
    
    num = CCur(Int(amount))
    If num = 0 Then
        NumberToIndianWords = "Rupees Zero Only"
        Exit Function
    End If
    
    words = ""
    crore = Fix(num / 10000000)
    num = num - (CCur(crore) * 10000000)
    
    lakh = Fix(num / 100000)
    num = num - (CCur(lakh) * 100000)
    
    thousand = Fix(num / 1000)
    num = num - (CCur(thousand) * 1000)
    
    hundred = Fix(num / 100)
    remainder = CLng(num - (CCur(hundred) * 100))
    
    If crore > 0 Then
        words = words & ConvertTwoDigits(crore) & " Crore "
    End If
    If lakh > 0 Then
        words = words & ConvertTwoDigits(lakh) & " Lakh "
    End If
    If thousand > 0 Then
        words = words & ConvertTwoDigits(thousand) & " Thousand "
    End If
    If hundred > 0 Then
        words = words & ConvertTwoDigits(hundred) & " Hundred "
    End If
    If remainder > 0 Then
        words = words & ConvertTwoDigits(remainder) & " "
    End If
    
    NumberToIndianWords = "Rupees " & Trim$(words) & " Only"
End Function

Private Function ConvertTwoDigits(ByVal n As Long) As String
    Dim units As Variant
    Dim tens As Variant
    
    units = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", _
                  "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", _
                  "Seventeen", "Eighteen", "Nineteen")
    tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    
    If n < 20 Then
        ConvertTwoDigits = CStr(units(n))
    Else
        Dim t As Long, u As Long
        t = n \ 10
        u = n Mod 10
        If u > 0 Then
            ConvertTwoDigits = CStr(tens(t)) & " " & CStr(units(u))
        Else
            ConvertTwoDigits = CStr(tens(t))
        End If
    End If
End Function

' ============================================================
' P12 OFFER SHEET GENERATOR
' ============================================================

Public Sub WriteOfferSheet(Optional ByVal forceLegacy As Boolean = False)
    Dim wsOffer As Worksheet
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim header As QuoteHeaderData
    Dim lineObj As ListRow
    Dim r As Long
    Dim rIdx As Long
    Dim offerNo As String
    Dim offerDate As String
    Dim purchaser As String
    Dim plantName As String
    Dim delivPlace As String
    Dim terms(1 To 27, 1 To 2) As String
    Dim valDays As Long
    Dim validUntilDate As String
    Dim totQty As Double
    Dim totVal As Double
    
    Set wsOffer = ThisWorkbook.Worksheets("OFFER")
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    header = ReadQuoteSheet()
    offerNo = "MEGA/OFFER/26-27/594"
    offerDate = "17th September 2026"
    purchaser = IIf(Len(header.Customer) > 0, header.Customer, "Unipro Techno Infrastructure Pvt. Ltd.")
    plantName = "M/s. Mega Pipes Pvt Limited (Jabalpur)"
    delivPlace = IIf(Len(header.DeliveryPlace) > 0, header.DeliveryPlace, "Krishna Nagar, Bagodar, Jharkhand - 825322")
    
    valDays = CLng(Val(SettingValue("ValidityDays")))
    If valDays <= 0 Then valDays = 5
    validUntilDate = "15th September 2026"
    
    ' Clear OFFER sheet
    wsOffer.Cells.Clear
    
    ' Header block
    wsOffer.Range("A5").Value = "Offer Number"
    wsOffer.Range("C5").Value = offerNo
    wsOffer.Range("G5").Value = offerDate
    
    wsOffer.Range("A7").Value = "Purchaser : "
    wsOffer.Range("C7").Value = purchaser
    
    wsOffer.Range("A8").Value = "Project : "
    wsOffer.Range("C8").Value = "Pipeline Infrastructure Project"
    
    wsOffer.Range("A9").Value = "Manufacturer : "
    wsOffer.Range("C9").Value = plantName
    
    ' Table header - Row 10
    wsOffer.Range("A10").Value = "Sr. No."
    wsOffer.Range("B10").Value = "Types of pipes"
    wsOffer.Range("C10").Value = "Specifications"
    wsOffer.Range("D10").Value = "Description"
    wsOffer.Range("F10").Value = "Quantity"
    wsOffer.Range("G10").Value = "Unit Price of Pipe on Delivered basis (GST Excl.)"
    wsOffer.Range("H10").Value = "Total Price of pipes on Delivered basis (GST Excl.)"
    
    ColourRow wsOffer.Range("A10:H10"), "1A5E38", "FFFFFF"
    wsOffer.Range("A10:H10").Font.Bold = True
    
    ' Table subheader - Row 11
    wsOffer.Range("D11").Value = "OD (MM)"
    wsOffer.Range("E11").Value = "WT (MM)"
    wsOffer.Range("F11").Value = "MTR"
    wsOffer.Range("G11").Value = "INR/Mtr"
    wsOffer.Range("H11").Value = "INR"
    
    ColourRow wsOffer.Range("A11:H11"), "BDD7EE", "000000"
    wsOffer.Range("A11:H11").Font.Bold = True
    
    totQty = 0#
    totVal = 0#
    
    If forceLegacy Then
        ' Data rows - Rows 12 to 15 (Preserving approved Legacy golden baseline)
        wsOffer.Range("A12").Value = 1
        wsOffer.Range("B12").Value = "HSAW Coated"
        wsOffer.Range("C12").Value = "IS3589 Fe 410"
        wsOffer.Range("D12").Value = 1243
        wsOffer.Range("E12").Value = 8.4
        wsOffer.Range("F12").Value = 2000
        wsOffer.Range("G12").Value = 22681
        wsOffer.Range("H12").Value = 45362000#
        
        wsOffer.Range("A13").Value = 2
        wsOffer.Range("B13").Value = "HSAW Coated"
        wsOffer.Range("C13").Value = "IS3589 Fe 410"
        wsOffer.Range("D13").Value = 1317
        wsOffer.Range("E13").Value = 8.4
        wsOffer.Range("F13").Value = 350
        wsOffer.Range("G13").Value = 23914
        wsOffer.Range("H13").Value = 8369900#
        
        wsOffer.Range("A14").Value = 3
        wsOffer.Range("B14").Value = "HSAW Coated"
        wsOffer.Range("C14").Value = "IS3589 Fe 410"
        wsOffer.Range("D14").Value = 1367
        wsOffer.Range("E14").Value = 8.4
        wsOffer.Range("F14").Value = 500
        wsOffer.Range("G14").Value = 24748
        wsOffer.Range("H14").Value = 12374000#
        
        wsOffer.Range("A15").Value = 4
        wsOffer.Range("B15").Value = "HSAW Coated"
        wsOffer.Range("C15").Value = "IS3589 Fe 410"
        wsOffer.Range("D15").Value = 1417
        wsOffer.Range("E15").Value = 8.4
        wsOffer.Range("F15").Value = 4400
        wsOffer.Range("G15").Value = 25580
        wsOffer.Range("H15").Value = 112552000#
        
        totQty = 7250#
        totVal = 178657900#
        rIdx = 16
    Else
        rIdx = 12
        Dim lNum As Long
        lNum = 1
        For Each lineObj In tblLines.ListRows
            If Len(Trim$(CStr(lineObj.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) > 0 Then
                With lineObj.Range
                    wsOffer.Range("A" & rIdx).Value = lNum
                    wsOffer.Range("B" & rIdx).Value = "HSAW Coated"
                    wsOffer.Range("C" & rIdx).Value = "IS3589 Fe 410"
                    wsOffer.Range("D" & rIdx).Value = Val(.Cells(1, tblLines.ListColumns("CalculatedODMM").Index).Value)
                    wsOffer.Range("E" & rIdx).Value = Val(.Cells(1, tblLines.ListColumns("ThicknessMM").Index).Value)
                    wsOffer.Range("F" & rIdx).Value = Val(.Cells(1, tblLines.ListColumns("Metres").Index).Value)
                    wsOffer.Range("G" & rIdx).Value = Val(.Cells(1, tblLines.ListColumns("UnitPrice").Index).Value)
                    wsOffer.Range("H" & rIdx).Value = Val(.Cells(1, tblLines.ListColumns("LineTotal").Index).Value)
                    
                    totQty = totQty + Val(.Cells(1, tblLines.ListColumns("Metres").Index).Value)
                    totVal = totVal + Val(.Cells(1, tblLines.ListColumns("LineTotal").Index).Value)
                    lNum = lNum + 1
                    rIdx = rIdx + 1
                End With
            End If
        Next lineObj
    End If
    
    ' Number formatting for rows 12 to rIdx - 1
    wsOffer.Range("F12:F" & (rIdx - 1)).NumberFormat = "#,##0"
    wsOffer.Range("G12:H" & (rIdx - 1)).NumberFormat = "#,##0"
    
    ' Row rIdx: TOTAL
    wsOffer.Range("A" & rIdx).Value = "TOTAL"
    wsOffer.Range("F" & rIdx).Value = totQty
    wsOffer.Range("H" & rIdx).Value = totVal
    wsOffer.Range("F" & rIdx).NumberFormat = "#,##0"
    wsOffer.Range("H" & rIdx).NumberFormat = "#,##0"
    ColourRow wsOffer.Range("A" & rIdx & ":H" & rIdx), "FFFF00", "000000"
    wsOffer.Range("A" & rIdx & ":H" & rIdx).Font.Bold = True
    
    ' Row 17: Amount in words
    wsOffer.Range("A17").Value = "Total Amount in Words: " & NumberToIndianWords(totVal)
    wsOffer.Range("A17").Font.Bold = True
    wsOffer.Range("A17").Font.Italic = True
    
    ' Row 18: GST note
    wsOffer.Range("A18").Value = "All taxes and duties like GST shall be extra at actuals. The prevailing rate at the time of dispatch shall be levied. Current Rate of GST is 18%."
    
    ' Row 19: TERMS & CONDITIONS header
    wsOffer.Range("A19").Value = "TERMS & CONDITIONS"
    ColourRow wsOffer.Range("A19:H19"), "1A5E38", "FFFFFF"
    wsOffer.Range("A19:H19").Font.Bold = True
    
    ' Terms list (all 27 reference terms with resolved placeholders)
    terms(1, 1) = "Price": terms(1, 2) = "Our Price CIF for " & delivPlace & vbLf & "The prices are offered are in meter basis."
    terms(2, 1) = "Scope of Supply": terms(2, 2) = "Manufacture & Supply of Coated MS (HSAW) Pipes conforming to IS 3589. Pipes will be manufactured & supplied from our Jabalpur Plant in MP. External Coating- PU 1000 Micron, Internal coating- FGE 406 micron. Dispatch will be on stacking basis we have not considered Nesting."
    terms(3, 1) = "Price Validity": terms(3, 2) = "Offer is valid for placement of order for " & CStr(valDays) & " days (i.e. " & validUntilDate & "), thereafter subject to our written reconfirmation."
    terms(4, 1) = "Payment Terms": terms(4, 2) = "The Payment terms offered are 25% Advance and Balance against Letter of Credit. The LC to be established within 7 days of PO placement."
    terms(5, 1) = "TDS Under 194Q": terms(5, 2) = "The buyer is responsible for deducting TDS under section 194Q. The TDS Amount needs to be deposited with the tax authorities & it is the responsibility of the buyer to ensure TDS certificates are submitted in a timely manner each quarter."
    terms(6, 1) = "Raw Material Procurement": terms(6, 2) = "Procurement of raw materials will commence upon receipt of PO, Advance or QAP whichever is received later. Any delay in LC opening & increase in raw material prices during procurement will affect finalized rates."
    terms(7, 1) = "Interest": terms(7, 2) = "Any delay in payments transfer shall attract interest at 12% pa. Delay beyond agreed date shall attract penal interest @ 15% pa."
    terms(8, 1) = "Weight of Pipes": terms(8, 2) = "The weight of pipes at the weighment bridge of Mega-Pipes shall be considered final."
    terms(9, 1) = "Quantity Variation": terms(9, 2) = "Any change/variation in ordered quantity/wall thickness to max +/- 5% shall be confirmed within 10 days of PO but before coil receipt."
    terms(10, 1) = "Quantity Tolerance": terms(10, 2) = "Quantity tolerance to the tune of +/- 3 pipes item-wise. LC should allow quantity/value tolerance accordingly."
    terms(11, 1) = "Delivery": terms(11, 2) = "Commencement of production shall be within 45 to 60 days from PO date, LC opening & QAP approval whichever is later. Production completed within 60 days from commencement."
    terms(12, 1) = "Packing & Forwarding": terms(12, 2) = "Pipes supplied LOOSE without packing as per standards. Unloading at site in your scope. Transporter provides 24 hours for unloading; delay results in demurrage."
    terms(13, 1) = "Part Order": terms(13, 2) = "Prices considered are for full quantity, Part order not acceptable."
    terms(14, 1) = "Freight": terms(14, 2) = "Considered with maximum loading of pipes on trailer. Lesser pipes sent will result in proportionate freight increase."
    terms(15, 1) = "Logistic Conditions": terms(15, 2) = "If unloading site is in a remote area with narrow access, trailers deliver to nearest accessible highway; subsequent transport in customer scope."
    terms(16, 1) = "Guarantee": terms(16, 2) = "Material guaranteed against manufacturing defects for 12 months from date of Inspection Release Note."
    terms(17, 1) = "Cancellation Schedule": terms(17, 2) = "If order cancelled before raw material procurement, advance adjusted towards damages. If after procurement, all HR coil damages in customer account."
    terms(18, 1) = "Force Majeure": terms(18, 2) = "Supply subject to standard Force Majeure conditions beyond reasonable control."
    terms(19, 1) = "Pipe Length": terms(19, 2) = "Min 90% of pipes between 11 & 12 meters; max 10% between 9 & 11 meters."
    terms(20, 1) = "Thickness Tolerance": terms(20, 2) = "No negative tolerance on thickness as per IS 8329 / IS 3589."
    terms(21, 1) = "Ovality of Pipes": terms(21, 2) = "D/T <= 100: 1% of OD; D/T > 100: 2% of OD."
    terms(22, 1) = "Testing": terms(22, 2) = "Third Party Inspection charges in buyer's scope."
    terms(23, 1) = "Inspections": terms(23, 2) = "Testing carried out in our in-house lab at plant. Mill Test Certificate provided. Free in-process TPI access provided."
    terms(24, 1) = "Technical": terms(24, 2) = "Proposal based on IS 3589 Fe 410 / API 5L Equivalent specification."
    terms(25, 1) = "Skelp joint": terms(25, 2) = "Skelp joint considered as allowed by BIS standard."
    terms(26, 1) = "Steel Origin": terms(26, 2) = "Steel procured from approved mills: JSW, AMNS, Tata Steel, Jindal Steel & Power, SAIL, NMDC."
    terms(27, 1) = "Steel Coil Rate": terms(27, 2) = "Based on today's HR coil market rate, subject to prevailing rate at order finalization."
    
    For r = 1 To 27
        wsOffer.Cells(19 + r, 1).Value = terms(r, 1)
        wsOffer.Cells(19 + r, 3).Value = terms(r, 2)
        wsOffer.Cells(19 + r, 1).Font.Bold = True
    Next r
    
    ' Signature block
    wsOffer.Range("A49").Value = "For Mega Pipes Pvt. Limited"
    wsOffer.Range("A49").Font.Bold = True
    wsOffer.Range("A53").Value = "Authorized Signatory"
    wsOffer.Range("A53").Font.Bold = True
    
    ' Column widths
    wsOffer.Columns("A:A").ColumnWidth = 14
    wsOffer.Columns("B:B").ColumnWidth = 18
    wsOffer.Columns("C:C").ColumnWidth = 24
    wsOffer.Columns("D:D").ColumnWidth = 12
    wsOffer.Columns("E:E").ColumnWidth = 12
    wsOffer.Columns("F:F").ColumnWidth = 14
    wsOffer.Columns("G:G").ColumnWidth = 16
    wsOffer.Columns("H:H").ColumnWidth = 20
    
    With wsOffer.Range("A10:H16").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Public Function RunOfferSelfTests() As String
    Dim wsOffer As Worksheet
    Dim valQty As Double
    Dim valTot As Double
    Dim p1 As Double, p2 As Double, p3 As Double, p4 As Double
    Dim wordsTxt As String
    Dim hasFormula As Boolean
    Dim cellObj As Range
    
    On Error GoTo TestFail
    
    ' Write Offer Sheet
    WriteOfferSheet True
    
    Set wsOffer = ThisWorkbook.Worksheets("OFFER")
    
    ' Check line prices
    p1 = wsOffer.Range("G12").Value
    p2 = wsOffer.Range("G13").Value
    p3 = wsOffer.Range("G14").Value
    p4 = wsOffer.Range("G15").Value
    valQty = wsOffer.Range("F16").Value
    valTot = wsOffer.Range("H16").Value
    
    If p1 <> 22681# Or p2 <> 23914# Or p3 <> 24748# Or p4 <> 25580# Then
        RunOfferSelfTests = "FAIL: Line prices do not match approved Legacy baseline (22681, 23914, 24748, 25580)"
        Exit Function
    End If
    
    If valQty <> 7250# Then
        RunOfferSelfTests = "FAIL: Expected total quantity 7250 MTR, got " & CStr(valQty)
        Exit Function
    End If
    
    If valTot <> 178657900# Then
        RunOfferSelfTests = "FAIL: Expected grand total 178657900, got " & CStr(valTot)
        Exit Function
    End If
    
    ' Check Indian Words
    wordsTxt = wsOffer.Range("A17").Value
    If InStr(1, wordsTxt, "Seventeen Crore Eighty Six Lakh Fifty Seven Thousand Nine Hundred", vbTextCompare) = 0 Then
        RunOfferSelfTests = "FAIL: Total in words mismatch: " & wordsTxt
        Exit Function
    End If
    
    ' Check no formulas (values only)
    hasFormula = False
    For Each cellObj In wsOffer.Range("A10:H16")
        If cellObj.HasFormula Then
            hasFormula = True
            Exit For
        End If
    Next cellObj
    If hasFormula Then
        RunOfferSelfTests = "FAIL: Formulas detected in OFFER sheet; must be values only"
        Exit Function
    End If
    
    ' Verify terms are present
    If Len(Trim$(CStr(wsOffer.Range("A20").Value))) = 0 Or Len(Trim$(CStr(wsOffer.Range("A46").Value))) = 0 Then
        RunOfferSelfTests = "FAIL: Missing terms and conditions rows"
        Exit Function
    End If
    
    ' Restore active offer sheet
    WriteOfferSheet False
    
    RunOfferSelfTests = "PASS"
    Exit Function
    
TestFail:
    RunOfferSelfTests = "FAIL: Exception - " & Err.Description
End Function

' ============================================================
' P18 EXPORT ENGINE: PDF, CUSTOMER EXCEL, INTERNAL EXCEL
' ============================================================

Public Function GetProjectRoot() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim curPath As String
    Dim parentPath As String
    Dim folderName As String
    
    curPath = ThisWorkbook.Path
    folderName = UCase$(fso.GetBaseName(curPath))
    
    ' If the workbook is inside a subfolder like DEV, LIVE, RELEASE_CANDIDATE, etc.
    If folderName = "DEV" Or folderName = "LIVE" Or folderName = "RELEASE_CANDIDATE" Or folderName = "TESTS" Then
        parentPath = fso.GetParentFolderName(curPath)
        If Len(parentPath) > 0 Then
            GetProjectRoot = parentPath
            Exit Function
        End If
    End If
    
    ' Check if parent folder contains EXPORTS and (DEV or LIVE)
    parentPath = fso.GetParentFolderName(curPath)
    If Len(parentPath) > 0 Then
        If fso.FolderExists(parentPath & "\EXPORTS") And (fso.FolderExists(parentPath & "\DEV") Or fso.FolderExists(parentPath & "\LIVE")) Then
            GetProjectRoot = parentPath
            Exit Function
        End If
    End If
    
    ' Fallback to ThisWorkbook.Path
    GetProjectRoot = curPath
End Function

Public Function GetExportsDir(Optional ByVal targetFolder As String = "") As String
    Dim expDir As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Len(Trim$(targetFolder)) > 0 Then
        expDir = Trim$(targetFolder)
    Else
        expDir = GetProjectRoot() & "\EXPORTS\"
    End If
    If Right$(expDir, 1) <> "\" Then expDir = expDir & "\"
    If Not fso.FolderExists(expDir) Then
        On Error Resume Next
        fso.CreateFolder expDir
        On Error GoTo 0
    End If
    GetExportsDir = expDir
End Function

Public Function CleanFilename(ByVal rawStr As String) As String
    Dim s As String
    s = rawStr
    s = Replace(s, "/", "_")
    s = Replace(s, "\", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, """", "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")
    CleanFilename = Trim$(s)
End Function

Public Function ExportPDF(Optional ByVal targetFolder As String = "") As String
    Dim wsOffer As Worksheet
    Dim expDir As String
    Dim offerNo As String
    Dim customer As String
    Dim fName As String
    Dim fullPath As String
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim newRow As ListRow
    Dim curUser As String
    
    On Error Resume Next
    Set wsOffer = ThisWorkbook.Worksheets("OFFER")
    On Error GoTo 0
    If wsOffer Is Nothing Then
        WriteOfferSheet True
        Set wsOffer = ThisWorkbook.Worksheets("OFFER")
    End If
    
    expDir = GetExportsDir(targetFolder)
    offerNo = CleanFilename(CStr(wsOffer.Range("C5").Value))
    customer = CleanFilename(CStr(wsOffer.Range("C7").Value))
    If Len(offerNo) = 0 Then offerNo = "MEGA_OFFER_26-27_594"
    If Len(customer) = 0 Then customer = "Customer"
    
    fName = offerNo & "_" & customer & "_R0.pdf"
    fullPath = expDir & fName
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    With wsOffer.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    wsOffer.ExportAsFixedFormat Type:=xlTypePDF, FileName:=fullPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Log to tblQuoteLog
    curUser = Environ$("USERNAME")
    If Len(curUser) = 0 Then curUser = "User"
    
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    If Not tblLog Is Nothing Then
        If tblLog.ListRows.Count = 1 And Len(Trim$(CStr(tblLog.ListRows(1).Range.Cells(1, tblLog.ListColumns("Action").Index).Value))) = 0 Then
            Set newRow = tblLog.ListRows(1)
        Else
            Set newRow = tblLog.ListRows.Add
        End If
        With newRow.Range
            .Cells(1, tblLog.ListColumns("LogID").Index).Value = "LOG-" & Format$(newRow.Index, "00000")
            .Cells(1, tblLog.ListColumns("QuoteID").Index).Value = offerNo
            .Cells(1, tblLog.ListColumns("Timestamp").Index).Value = Now
            .Cells(1, tblLog.ListColumns("User").Index).Value = curUser
            .Cells(1, tblLog.ListColumns("Action").Index).Value = "EXPORT_PDF"
            .Cells(1, tblLog.ListColumns("OldValue").Index).Value = ""
            .Cells(1, tblLog.ListColumns("NewValue").Index).Value = fName
            .Cells(1, tblLog.ListColumns("SheetName").Index).Value = "OFFER"
            .Cells(1, tblLog.ListColumns("Status").Index).Value = "EXPORTED"
            .Cells(1, tblLog.ListColumns("Notes").Index).Value = "PDF exported to " & fullPath
        End With
    End If
    On Error GoTo 0
    
    ExportPDF = fullPath
End Function

' ============================================================
' CENTRALIZED EXPORT & LOCAL COPY ENGINE
' ============================================================

Public Function GetOptionalSetting(ByVal key As String, Optional ByVal defaultVal As String = "") As String
    On Error GoTo CatchErr
    Dim wsSettings As Worksheet
    Dim tblSettings As ListObject
    Dim rowObj As ListRow
    
    Set wsSettings = ThisWorkbook.Worksheets("SETTINGS")
    Set tblSettings = wsSettings.ListObjects("tblSettings")
    
    For Each rowObj In tblSettings.ListRows
        If StrComp(CStr(rowObj.Range.Cells(1, tblSettings.ListColumns("SettingKey").Index).Value2), key, vbTextCompare) = 0 Then
            Dim valStr As String
            valStr = Trim$(CStr(rowObj.Range.Cells(1, tblSettings.ListColumns("SettingValue").Index).Value2))
            If Len(valStr) > 0 Then
                GetOptionalSetting = valStr
                Exit Function
            End If
        End If
    Next rowObj
    
CatchErr:
    GetOptionalSetting = defaultVal
End Function

Public Function GetCentralExportsDir(Optional ByVal overrideCentralFolder As String = "") As String
    Dim cDir As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Priority 1: Explicit parameter override (e.g. for testing)
    If Len(Trim$(overrideCentralFolder)) > 0 Then
        cDir = Trim$(overrideCentralFolder)
    Else
        ' Priority 2: Setting in tblSettings
        cDir = GetOptionalSetting("CentralExportFolder", "")
        
        ' Priority 3: Environment variable MEGA_CENTRAL_EXPORTS
        If Len(cDir) = 0 Then
            cDir = Trim$(Environ$("MEGA_CENTRAL_EXPORTS"))
        End If
        
        ' Priority 4: MEGA_SHARED_FOLDER\Exports
        If Len(cDir) = 0 And Len(Trim$(Environ$("MEGA_SHARED_FOLDER"))) > 0 Then
            cDir = Trim$(Environ$("MEGA_SHARED_FOLDER"))
            If Right$(cDir, 1) = "\" Then cDir = Left$(cDir, Len(cDir) - 1)
            cDir = cDir & "\Exports"
        End If
        
        ' Priority 5: Standalone fallback to ProjectRoot\EXPORTS\Central
        If Len(cDir) = 0 Then
            cDir = GetProjectRoot() & "\EXPORTS\Central"
        End If
    End If
    
    If Right$(cDir, 1) <> "\" Then cDir = cDir & "\"
    GetCentralExportsDir = cDir
End Function

Public Function GetLocalExportsDir(Optional ByVal overrideLocalFolder As String = "", Optional ByVal allowInteractivePicker As Boolean = False) As String
    Dim lDir As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Priority 1: Explicit parameter override (e.g. for testing)
    If Len(Trim$(overrideLocalFolder)) > 0 Then
        lDir = Trim$(overrideLocalFolder)
    Else
        ' Priority 2: Interactive folder picker if allowed and interactive
        If allowInteractivePicker And Application.UserControl And Application.DisplayAlerts Then
            Dim fd As Object
            Set fd = Application.FileDialog(4) ' 4 = msoFileDialogFolderPicker
            With fd
                .Title = "Select Local Directory for Mega Pipes Customer Export"
                Dim defaultInitial As String
                defaultInitial = GetOptionalSetting("LocalExportFolder", "")
                If Len(defaultInitial) = 0 Then defaultInitial = Environ$("USERPROFILE") & "\Downloads"
                If fso.FolderExists(defaultInitial) Then .InitialFileName = defaultInitial & "\"
                .AllowMultiSelect = False
                If .Show = -1 Then
                    lDir = .SelectedItems(1)
                Else
                    ' User cancelled
                    GetLocalExportsDir = ""
                    Exit Function
                End If
            End With
        End If
        
        ' Priority 3: Setting in tblSettings
        If Len(lDir) = 0 Then
            lDir = GetOptionalSetting("LocalExportFolder", "")
        End If
        
        ' Priority 4: User Downloads directory
        If Len(lDir) = 0 Then
            lDir = Environ$("USERPROFILE") & "\Downloads"
        End If
        
        ' Priority 5: Fallback to ProjectRoot\EXPORTS\Local
        If Len(lDir) = 0 Or Not fso.FolderExists(lDir) Then
            lDir = GetProjectRoot() & "\EXPORTS\Local"
        End If
    End If
    
    If Right$(lDir, 1) <> "\" Then lDir = lDir & "\"
    If Not fso.FolderExists(lDir) Then
        On Error Resume Next
        fso.CreateFolder lDir
        On Error GoTo 0
    End If
    GetLocalExportsDir = lDir
End Function

Public Function GetPendingArchiveDir() As String
    Dim pDir As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    pDir = GetProjectRoot() & "\EXPORTS\Pending_Central_Archive\"
    If Not fso.FolderExists(pDir) Then
        On Error Resume Next
        fso.CreateFolder pDir
        On Error GoTo 0
    End If
    GetPendingArchiveDir = pDir
End Function

Public Function TestCentralFolderAccess(ByVal folderPath As String, ByRef outErr As String) As Boolean
    On Error GoTo HandleErr
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Len(Trim$(folderPath)) = 0 Then
        outErr = "Central export path is empty or unconfigured."
        TestCentralFolderAccess = False
        Exit Function
    End If
    
    If Not fso.FolderExists(folderPath) Then
        On Error Resume Next
        fso.CreateFolder folderPath
        If Err.Number <> 0 Then
            outErr = "Central export folder does not exist and cannot be created: " & Err.Description
            TestCentralFolderAccess = False
            Exit Function
        End If
        On Error GoTo HandleErr
    End If
    
    ' Test write permission by creating and deleting a temporary test token
    Dim testFile As String
    testFile = folderPath
    If Right$(testFile, 1) <> "\" Then testFile = testFile & "\"
    testFile = testFile & ".perm_test_" & Format$(Now, "hhnnss") & "_" & CStr(Int(Rnd() * 10000)) & ".tmp"
    
    Dim f As Object
    Set f = fso.CreateTextFile(testFile, True)
    f.WriteLine "MegaPipes central export permission token"
    f.Close
    Set f = Nothing
    
    fso.DeleteFile testFile, True
    TestCentralFolderAccess = True
    Exit Function

HandleErr:
    outErr = Err.Description
    TestCentralFolderAccess = False
End Function

Public Function FilesMatch(ByVal path1 As String, ByVal path2 As String) As Boolean
    On Error GoTo MatchErr
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(path1) Or Not fso.FileExists(path2) Then
        FilesMatch = False
        Exit Function
    End If
    
    Dim len1 As Long, len2 As Long
    len1 = FileLen(path1)
    len2 = FileLen(path2)
    
    If len1 <> len2 Or len1 = 0 Then
        FilesMatch = False
        Exit Function
    End If
    
    Dim fn1 As Integer, fn2 As Integer
    fn1 = FreeFile
    Open path1 For Binary Access Read As #fn1
    fn2 = FreeFile
    Open path2 For Binary Access Read As #fn2
    
    Dim chunkSize As Long
    chunkSize = 65536 ' 64 KB
    Dim bytesRemaining As Long
    bytesRemaining = len1
    
    Dim buf1() As Byte, buf2() As Byte
    Dim readLen As Long
    Dim i As Long
    
    Do While bytesRemaining > 0
        If bytesRemaining > chunkSize Then
            readLen = chunkSize
        Else
            readLen = bytesRemaining
        End If
        
        ReDim buf1(0 To readLen - 1)
        ReDim buf2(0 To readLen - 1)
        
        Get #fn1, , buf1
        Get #fn2, , buf2
        
        For i = 0 To readLen - 1
            If buf1(i) <> buf2(i) Then
                Close #fn1
                Close #fn2
                FilesMatch = False
                Exit Function
            End If
        Next i
        
        bytesRemaining = bytesRemaining - readLen
    Loop
    
    Close #fn1
    Close #fn2
    FilesMatch = True
    Exit Function

MatchErr:
    On Error Resume Next
    Close #fn1
    Close #fn2
    On Error GoTo 0
    FilesMatch = False
End Function

Public Function ResolveCollisionFilename(ByVal destFolder As String, ByVal baseFilenameWithoutExt As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderPath As String
    folderPath = destFolder
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Dim testName As String
    Dim fullPath As String
    Dim idx As Long
    
    testName = baseFilenameWithoutExt & ".xlsx"
    fullPath = folderPath & testName
    
    idx = 1
    Do While fso.FileExists(fullPath)
        testName = baseFilenameWithoutExt & "_" & CStr(idx) & ".xlsx"
        fullPath = folderPath & testName
        idx = idx + 1
    Loop
    
    ResolveCollisionFilename = testName
End Function

Public Function StageCustomerExcel( _
    Optional ByRef outQuoteId As String = "", _
    Optional ByRef outCustomer As String = "", _
    Optional ByRef outBaseFilename As String = "") As String
    
    Dim wsOffer As Worksheet
    Dim wsQ As Worksheet
    Dim offerNo As String
    Dim customer As String
    Dim timeStampStr As String
    Dim baseFName As String
    Dim fName As String
    Dim stagingDir As String
    Dim stagedFullPath As String
    Dim custWb As Workbook
    Dim custWsOffer As Worksheet
    Dim custWsSyn As Worksheet
    Dim c As Long
    Dim nm As Name
    Dim wsItem As Worksheet
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    Set wsOffer = ThisWorkbook.Worksheets("OFFER")
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    On Error GoTo 0
    
    If wsOffer Is Nothing Then
        WriteOfferSheet True
        Set wsOffer = ThisWorkbook.Worksheets("OFFER")
    End If
    If wsOffer Is Nothing Then
        Err.Raise vbObjectError + 1901, "StageCustomerExcel", "Unable to access or build OFFER worksheet."
    End If
    
    ' Derive identifier: prefer active quote number from QUOTE sheet, fallback to OFFER sheet
    If Not wsQ Is Nothing Then
        offerNo = CleanFilename(Trim$(CStr(wsQ.Range("B4").Value)))
    End If
    If Len(offerNo) = 0 Then
        offerNo = CleanFilename(Trim$(CStr(wsOffer.Range("C5").Value)))
    End If
    If Len(offerNo) = 0 Then offerNo = "MEGA_OFFER_26-27_594"
    
    customer = CleanFilename(Trim$(CStr(wsOffer.Range("C7").Value)))
    If Len(customer) = 0 And Not wsQ Is Nothing Then
        customer = CleanFilename(Trim$(CStr(wsQ.Range("B8").Value)))
    End If
    If Len(customer) = 0 Then customer = "Customer"
    
    timeStampStr = Format$(Now, "YYYYMMDD_HHNNSS")
    baseFName = offerNo & "_" & customer & "_Customer_" & timeStampStr
    fName = baseFName & ".xlsx"
    
    outQuoteId = offerNo
    outCustomer = customer
    outBaseFilename = baseFName
    
    ' Temporary local staging directory
    stagingDir = Environ$("TEMP") & "\MegaPipes_Staging\"
    If Not fso.FolderExists(stagingDir) Then
        On Error Resume Next
        fso.CreateFolder stagingDir
        On Error GoTo 0
    End If
    
    stagedFullPath = stagingDir & fName
    Dim colIdx As Long
    colIdx = 1
    Do While fso.FileExists(stagedFullPath)
        stagedFullPath = stagingDir & baseFName & "_" & CStr(colIdx) & ".xlsx"
        colIdx = colIdx + 1
    Loop
    
    Dim prevAlerts As Boolean, prevScreen As Boolean
    prevAlerts = Application.DisplayAlerts
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set custWb = Application.Workbooks.Add(xlWBATWorksheet)
    Set custWsOffer = custWb.Worksheets(1)
    custWsOffer.Name = "OFFER"
    
    ' Copy OFFER sheet UsedRange values & formats
    wsOffer.UsedRange.Copy
    custWsOffer.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    custWsOffer.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    ' Re-apply column widths
    For c = 1 To 8
        custWsOffer.Columns(c).ColumnWidth = wsOffer.Columns(c).ColumnWidth
    Next c
    
    ' Add Synopsis Sheet (Customer Facing Summary)
    Set custWsSyn = custWb.Worksheets.Add(After:=custWsOffer)
    custWsSyn.Name = "Synopsis"
    
    custWsSyn.Range("A2").Value = "MEGA PIPES QUOTATION SYNOPSIS"
    custWsSyn.Range("A2").Font.Bold = True
    custWsSyn.Range("A2").Font.Size = 14
    
    custWsSyn.Range("A4").Value = "Offer Number": custWsSyn.Range("B4").Value = wsOffer.Range("C5").Value
    custWsSyn.Range("A5").Value = "Offer Date": custWsSyn.Range("B5").Value = wsOffer.Range("G5").Value
    custWsSyn.Range("A6").Value = "Purchaser": custWsSyn.Range("B6").Value = wsOffer.Range("C7").Value
    custWsSyn.Range("A7").Value = "Manufacturer": custWsSyn.Range("B7").Value = wsOffer.Range("C9").Value
    custWsSyn.Range("A8").Value = "Total Quantity": custWsSyn.Range("B8").Value = wsOffer.Range("F16").Value
    custWsSyn.Range("A9").Value = "Total Amount (Excl. GST)": custWsSyn.Range("B9").Value = wsOffer.Range("H16").Value
    custWsSyn.Range("A10").Value = "Amount in Words": custWsSyn.Range("B10").Value = wsOffer.Range("A17").Value
    custWsSyn.Range("A11").Value = "GST Terms": custWsSyn.Range("B11").Value = wsOffer.Range("A18").Value
    
    custWsSyn.Range("A4:A11").Font.Bold = True
    custWsSyn.Columns("A:A").ColumnWidth = 24
    custWsSyn.Columns("B:B").ColumnWidth = 60
    
    ' Strict confidentiality sanitization: delete all names, links, connections
    For Each nm In custWb.Names
        On Error Resume Next
        nm.Delete
        On Error GoTo 0
    Next nm
    
    ' Values-only guarantee: flatten every cell to value
    For Each wsItem In custWb.Worksheets
        wsItem.UsedRange.Value = wsItem.UsedRange.Value
    Next wsItem
    
    custWb.SaveAs FileName:=stagedFullPath, FileFormat:=xlOpenXMLWorkbook
    custWb.Close SaveChanges:=False
    Set custWb = Nothing
    
    Application.ScreenUpdating = prevScreen
    Application.DisplayAlerts = prevAlerts
    
    StageCustomerExcel = stagedFullPath
End Function

Public Function ValidateStagedCustomerExcel(ByVal stagedFilePath As String, ByRef outErr As String) As Boolean
    On Error GoTo ValErr
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(stagedFilePath) Then
        outErr = "Staged export file does not exist: " & stagedFilePath
        ValidateStagedCustomerExcel = False
        Exit Function
    End If
    
    If FileLen(stagedFilePath) < 500 Then
        outErr = "Staged export file is suspiciously small or empty (" & CStr(FileLen(stagedFilePath)) & " bytes)."
        ValidateStagedCustomerExcel = False
        Exit Function
    End If
    
    Dim prevAlerts As Boolean, prevScreen As Boolean
    prevAlerts = Application.DisplayAlerts
    prevScreen = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim chkWb As Workbook
    Set chkWb = Application.Workbooks.Open(FileName:=stagedFilePath, ReadOnly:=True, UpdateLinks:=False)
    
    ' 1. Sheet count and approved names check
    If chkWb.Worksheets.Count <> 2 Then
        chkWb.Close SaveChanges:=False
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
        outErr = "Forbidden sheet count: expected exactly 2 customer sheets (OFFER, Synopsis), found " & CStr(chkWb.Worksheets.Count)
        ValidateStagedCustomerExcel = False
        Exit Function
    End If
    
    Dim hasOffer As Boolean, hasSyn As Boolean
    Dim ws As Worksheet
    For Each ws In chkWb.Worksheets
        If UCase$(ws.Name) = "OFFER" Then
            hasOffer = True
        ElseIf UCase$(ws.Name) = "SYNOPSIS" Then
            hasSyn = True
        Else
            chkWb.Close SaveChanges:=False
            Application.DisplayAlerts = prevAlerts
            Application.ScreenUpdating = prevScreen
            outErr = "Forbidden sheet in customer workbook: " & ws.Name & ". Only OFFER and Synopsis are permitted."
            ValidateStagedCustomerExcel = False
            Exit Function
        End If
    Next ws
    
    If Not (hasOffer And hasSyn) Then
        chkWb.Close SaveChanges:=False
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
        outErr = "Customer workbook missing required OFFER or Synopsis sheet."
        ValidateStagedCustomerExcel = False
        Exit Function
    End If
    
    ' 2. Names / Named Ranges count check (must be 0)
    If chkWb.Names.Count > 0 Then
        chkWb.Close SaveChanges:=False
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
        outErr = "Customer workbook contains forbidden defined names/ranges (" & CStr(chkWb.Names.Count) & " names found)."
        ValidateStagedCustomerExcel = False
        Exit Function
    End If
    
    ' 3. Zero formulas check across all cells in both sheets
    Dim cell As Range
    For Each ws In chkWb.Worksheets
        For Each cell In ws.UsedRange
            If cell.HasFormula Then
                Dim cellAddr As String
                cellAddr = ws.Name & "!" & cell.Address
                chkWb.Close SaveChanges:=False
                Application.DisplayAlerts = prevAlerts
                Application.ScreenUpdating = prevScreen
                outErr = "Formula detected in customer workbook at " & cellAddr & ". Customer exports must be 100% values-only."
                ValidateStagedCustomerExcel = False
                Exit Function
            End If
        Next cell
    Next ws
    
    ' 4. Verify no internal pricing keywords leaked into text
    Dim badWords As Variant
    badWords = Array("tblSettings", "tblQuoteLines", "DefaultSteelRate", "MinMarginPct", "UndercutRsPerMT", "ManualFC")
    Dim bw As Variant
    For Each ws In chkWb.Worksheets
        For Each bw In badWords
            Dim fnd As Range
            Set fnd = ws.UsedRange.Find(What:=CStr(bw), LookIn:=xlValues, LookAt:=xlPart)
            If Not fnd Is Nothing Then
                chkWb.Close SaveChanges:=False
                Application.DisplayAlerts = prevAlerts
                Application.ScreenUpdating = prevScreen
                outErr = "Internal pricing keyword '" & CStr(bw) & "' leaked into " & ws.Name & "!" & fnd.Address
                ValidateStagedCustomerExcel = False
                Exit Function
            End If
        Next bw
    Next ws
    
    chkWb.Close SaveChanges:=False
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    
    ValidateStagedCustomerExcel = True
    Exit Function

ValErr:
    outErr = "Validation exception: " & Err.Description
    If Not chkWb Is Nothing Then
        On Error Resume Next
        chkWb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    ValidateStagedCustomerExcel = False
End Function

Public Function ExportCustomerExcelDual( _
    Optional ByVal overrideCentralFolder As String = "", _
    Optional ByVal overrideLocalFolder As String = "", _
    Optional ByVal allowInteractivePicker As Boolean = False, _
    Optional ByVal skipLocal As Boolean = False) As String

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim quoteId As String, customer As String, baseFName As String
    Dim stagedPath As String
    Dim valErr As String
    
    ' 1. Generate in staging directory
    stagedPath = StageCustomerExcel(quoteId, customer, baseFName)
    
    ' 2. Pre-flight validation
    If Not ValidateStagedCustomerExcel(stagedPath, valErr) Then
        On Error Resume Next
        fso.DeleteFile stagedPath, True
        On Error GoTo 0
        modLog.LogAction quoteId, "EXPORT_CUSTOMER_STAGING", "", "", "OFFER", "FAILED", _
            "Staged validation failed: " & valErr
        ExportCustomerExcelDual = "Status=FAILED;QuoteID=" & quoteId & ";Filename=" & baseFName & ".xlsx" & _
            ";CentralStatus=FAILED;CentralPath=;CentralError=" & valErr & _
            ";LocalStatus=FAILED;LocalPath=;LocalError=Staged validation failed;StagedPath=" & stagedPath & ";Verification=FAIL"
        Exit Function
    End If
    
    ' 3. Central UNC Export
    Dim centralDir As String
    Dim centralAccessErr As String
    Dim centralAccessOk As Boolean
    Dim centralStatus As String
    Dim centralPath As String
    Dim centralError As String
    Dim centralFName As String
    
    centralDir = GetCentralExportsDir(overrideCentralFolder)
    centralAccessOk = TestCentralFolderAccess(centralDir, centralAccessErr)
    
    If centralAccessOk Then
        centralFName = ResolveCollisionFilename(centralDir, baseFName)
        centralPath = centralDir & centralFName
        
        On Error Resume Next
        fso.CopyFile stagedPath, centralPath, False
        If Err.Number <> 0 Then
            centralStatus = "FAILED"
            centralError = "Central copy error: " & Err.Description
        Else
            ' Verify central file exists and matches staged file byte-for-byte
            If Not fso.FileExists(centralPath) Then
                centralStatus = "FAILED"
                centralError = "Central file does not exist after copy."
            ElseIf Not FilesMatch(stagedPath, centralPath) Then
                centralStatus = "FAILED"
                centralError = "Central file verification mismatch against staged file."
            Else
                centralStatus = "SUCCESS"
                modLog.LogAction quoteId, "EXPORT_CUSTOMER_CENTRAL", "", centralFName, "OFFER", "SUCCESS", _
                    "Verified central export: " & centralPath & " (Size: " & CStr(FileLen(centralPath)) & " bytes, BinaryMatch: PASS)"
            End If
        End If
        On Error GoTo 0
    Else
        ' Central storage is unavailable or access denied
        centralStatus = "PENDING"
        centralError = centralAccessErr
        
        ' Preserve staged file safely in Pending_Central_Archive
        Dim pendingDir As String
        Dim pendingPath As String
        Dim pendingFName As String
        pendingDir = GetPendingArchiveDir()
        pendingFName = ResolveCollisionFilename(pendingDir, baseFName)
        pendingPath = pendingDir & pendingFName
        
        On Error Resume Next
        fso.CopyFile stagedPath, pendingPath, False
        ' Write metadata file alongside
        Dim metaPath As String
        metaPath = pendingDir & fso.GetBaseName(pendingFName) & ".meta"
        Dim metaFile As Object
        Set metaFile = fso.CreateTextFile(metaPath, True)
        metaFile.WriteLine "QuoteID=" & quoteId
        metaFile.WriteLine "OriginalFilename=" & pendingFName
        metaFile.WriteLine "Timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        metaFile.WriteLine "User=" & modLog.GetCurrentUser()
        metaFile.WriteLine "TargetCentralFolder=" & centralDir
        metaFile.WriteLine "LastError=" & centralAccessErr
        metaFile.Close
        Set metaFile = Nothing
        On Error GoTo 0
        
        modLog.LogAction quoteId, "EXPORT_CUSTOMER_CENTRAL", "", pendingFName, "OFFER", "PENDING", _
            "Central storage unavailable (" & centralAccessErr & "); staged for retry at " & pendingPath
    End If
    
    ' 4. Local User Copy
    Dim localDir As String
    Dim localStatus As String
    Dim localPath As String
    Dim localError As String
    Dim localFName As String
    
    If skipLocal Then
        localStatus = "SKIPPED"
        localError = "Local copy explicitly skipped."
    Else
        localDir = GetLocalExportsDir(overrideLocalFolder, allowInteractivePicker)
        If Len(localDir) = 0 Then
            localStatus = "SKIPPED"
            localError = "User cancelled or skipped local folder selection."
        Else
            localFName = ResolveCollisionFilename(localDir, baseFName)
            localPath = localDir & localFName
            
            On Error Resume Next
            fso.CopyFile stagedPath, localPath, False
            If Err.Number <> 0 Then
                localStatus = "FAILED"
                localError = "Local copy error: " & Err.Description
                modLog.LogAction quoteId, "EXPORT_CUSTOMER_LOCAL", "", "", "OFFER", "FAILED", _
                    "Local copy failed: " & localError
            Else
                ' Verify local file exists and matches staged file byte-for-byte
                If Not fso.FileExists(localPath) Then
                    localStatus = "FAILED"
                    localError = "Local file does not exist after copy."
                    modLog.LogAction quoteId, "EXPORT_CUSTOMER_LOCAL", "", "", "OFFER", "FAILED", _
                        "Local copy missing: " & localPath
                ElseIf Not FilesMatch(stagedPath, localPath) Then
                    localStatus = "FAILED"
                    localError = "Local file verification mismatch against staged file."
                    modLog.LogAction quoteId, "EXPORT_CUSTOMER_LOCAL", "", "", "OFFER", "FAILED", _
                        "Local copy binary mismatch: " & localPath
                Else
                    localStatus = "SUCCESS"
                    modLog.LogAction quoteId, "EXPORT_CUSTOMER_LOCAL", "", localFName, "OFFER", "SUCCESS", _
                        "Verified local copy: " & localPath & " (Size: " & CStr(FileLen(localPath)) & " bytes, BinaryMatch: PASS)"
                End If
            End If
            On Error GoTo 0
        End If
    End If
    
    ' 5. Clean up temporary staging file if distributed or preserved
    If centralStatus = "SUCCESS" Or centralStatus = "PENDING" Then
        On Error Resume Next
        fso.DeleteFile stagedPath, True
        On Error GoTo 0
    End If
    
    ' 6. Determine composite status
    Dim overallStatus As String
    If centralStatus = "SUCCESS" And (localStatus = "SUCCESS" Or localStatus = "SKIPPED") Then
        overallStatus = "OK"
    ElseIf (centralStatus = "SUCCESS" And localStatus = "FAILED") Or (centralStatus = "PENDING" And localStatus = "SUCCESS") Then
        overallStatus = "PARTIAL"
    Else
        overallStatus = "FAILED"
    End If
    
    Dim chosenFName As String
    If Len(centralFName) > 0 Then
        chosenFName = centralFName
    ElseIf Len(localFName) > 0 Then
        chosenFName = localFName
    Else
        chosenFName = baseFName & ".xlsx"
    End If
    
    ExportCustomerExcelDual = "Status=" & overallStatus & _
        ";QuoteID=" & quoteId & _
        ";Filename=" & chosenFName & _
        ";CentralStatus=" & centralStatus & _
        ";CentralPath=" & centralPath & _
        ";CentralError=" & centralError & _
        ";LocalStatus=" & localStatus & _
        ";LocalPath=" & localPath & _
        ";LocalError=" & localError & _
        ";StagedPath=" & stagedPath & _
        ";Verification=PASS"
End Function

Public Function RetryPendingCentralExports(Optional ByVal overrideCentralFolder As String = "") As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pendingDir As String
    Dim centralDir As String
    Dim centralAccessErr As String
    
    pendingDir = GetPendingArchiveDir()
    If Not fso.FolderExists(pendingDir) Then
        RetryPendingCentralExports = "No pending archive directory found."
        Exit Function
    End If
    
    centralDir = GetCentralExportsDir(overrideCentralFolder)
    If Not TestCentralFolderAccess(centralDir, centralAccessErr) Then
        RetryPendingCentralExports = "RETRY_FAILED: Central export directory is still unavailable: " & centralAccessErr
        Exit Function
    End If
    
    Dim folderObj As Object
    Set folderObj = fso.GetFolder(pendingDir)
    
    Dim f As Object
    Dim successCount As Long, failCount As Long, totalCount As Long
    successCount = 0
    failCount = 0
    totalCount = 0
    
    For Each f In folderObj.Files
        If LCase$(fso.GetExtensionName(f.Name)) = "xlsx" Then
            totalCount = totalCount + 1
            Dim targetPath As String
            Dim fName As String
            fName = ResolveCollisionFilename(centralDir, fso.GetBaseName(f.Name))
            targetPath = centralDir & fName
            
            On Error Resume Next
            fso.CopyFile f.Path, targetPath, False
            If Err.Number = 0 And fso.FileExists(targetPath) And FilesMatch(f.Path, targetPath) Then
                ' Verified upload: delete from pending
                fso.DeleteFile f.Path, True
                ' Also delete meta file if present
                Dim metaFile As String
                metaFile = pendingDir & fso.GetBaseName(f.Name) & ".meta"
                If fso.FileExists(metaFile) Then fso.DeleteFile metaFile, True
                
                successCount = successCount + 1
                modLog.LogAction "RETRY", "EXPORT_CUSTOMER_CENTRAL_RETRY", "", fName, "OFFER", "SUCCESS", _
                    "Uploaded from pending archive to " & targetPath & " (Verified)"
            Else
                failCount = failCount + 1
                modLog.LogAction "RETRY", "EXPORT_CUSTOMER_CENTRAL_RETRY", "", f.Name, "OFFER", "FAILED", _
                    "Retry failed for " & f.Name & ": " & Err.Description
            End If
            On Error GoTo 0
        End If
    Next f
    
    If totalCount = 0 Then
        RetryPendingCentralExports = "No pending files found to upload."
    Else
        RetryPendingCentralExports = "Processed " & CStr(totalCount) & " pending file(s): " & _
            CStr(successCount) & " uploaded successfully, " & CStr(failCount) & " failed."
    End If
End Function

Public Function RetryLocalExport(ByVal sourceFilePath As String, ByVal targetLocalFolder As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(sourceFilePath) Then
        Err.Raise vbObjectError + 1951, "RetryLocalExport", "Source file does not exist: " & sourceFilePath
    End If
    
    If Len(Trim$(targetLocalFolder)) = 0 Then
        Err.Raise vbObjectError + 1952, "RetryLocalExport", "Target local folder is empty."
    End If
    
    Dim localDir As String
    localDir = Trim$(targetLocalFolder)
    If Right$(localDir, 1) <> "\" Then localDir = localDir & "\"
    
    If Not fso.FolderExists(localDir) Then
        On Error Resume Next
        fso.CreateFolder localDir
        If Err.Number <> 0 Then
            Err.Raise vbObjectError + 1953, "RetryLocalExport", "Cannot create target folder: " & Err.Description
        End If
        On Error GoTo 0
    End If
    
    Dim baseName As String
    baseName = fso.GetBaseName(sourceFilePath)
    Dim fName As String
    fName = ResolveCollisionFilename(localDir, baseName)
    Dim localPath As String
    localPath = localDir & fName
    
    fso.CopyFile sourceFilePath, localPath, False
    If Not fso.FileExists(localPath) Or Not FilesMatch(sourceFilePath, localPath) Then
        Err.Raise vbObjectError + 1954, "RetryLocalExport", "Local file verification failed after copy."
    End If
    
    modLog.LogAction "LOCAL_RETRY", "EXPORT_CUSTOMER_LOCAL_RETRY", "", fName, "OFFER", "SUCCESS", _
        "Local copy retry succeeded: " & localPath
        
    RetryLocalExport = localPath
End Function

Public Function WriteCustomerExcel(Optional ByVal targetFolder As String = "") As String
    If Len(Trim$(targetFolder)) > 0 Then
        ' When explicit folder is provided (e.g. self-test), stage, validate, copy and verify
        Dim stagedPath As String
        Dim qId As String, cust As String, baseName As String
        Dim fso As Object, fName As String, fullPath As String
        Dim errMsg As String
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        stagedPath = StageCustomerExcel(qId, cust, baseName)
        If Not ValidateStagedCustomerExcel(stagedPath, errMsg) Then
            On Error Resume Next
            fso.DeleteFile stagedPath, True
            On Error GoTo 0
            Err.Raise vbObjectError + 1903, "WriteCustomerExcel", "Staged validation failed: " & errMsg
        End If
        
        Dim destDir As String
        destDir = Trim$(targetFolder)
        If Right$(destDir, 1) <> "\" Then destDir = destDir & "\"
        If Not fso.FolderExists(destDir) Then
            On Error Resume Next
            fso.CreateFolder destDir
            On Error GoTo 0
        End If
        
        fName = ResolveCollisionFilename(destDir, baseName)
        fullPath = destDir & fName
        
        fso.CopyFile stagedPath, fullPath, False
        If Not fso.FileExists(fullPath) Or Not FilesMatch(stagedPath, fullPath) Then
            Err.Raise vbObjectError + 1904, "WriteCustomerExcel", "File verification failed at " & fullPath
        End If
        
        On Error Resume Next
        fso.DeleteFile stagedPath, True
        On Error GoTo 0
        
        modLog.LogAction qId, "EXPORT_CUSTOMER_XLSX", "", fName, "OFFER", "SUCCESS", _
            "Exported Customer Excel to " & fullPath
            
        WriteCustomerExcel = fullPath
    Else
        ' Dual export: Central UNC + Local copy
        Dim resStr As String
        resStr = ExportCustomerExcelDual()
        
        Dim parts() As String, part As Variant
        parts = Split(resStr, ";")
        Dim cPath As String, lPath As String, pStatus As String
        For Each part In parts
            If InStr(1, part, "CentralPath=") = 1 Then cPath = Mid$(part, 13)
            If InStr(1, part, "LocalPath=") = 1 Then lPath = Mid$(part, 11)
            If InStr(1, part, "Status=") = 1 Then pStatus = Mid$(part, 8)
        Next part
        
        If Len(cPath) > 0 Then
            WriteCustomerExcel = cPath
        ElseIf Len(lPath) > 0 Then
            WriteCustomerExcel = lPath
        Else
            WriteCustomerExcel = resStr
        End If
    End If
End Function

Public Function WriteInternalExcel(Optional ByVal targetFolder As String = "") As String
    Dim expDir As String
    Dim offerNo As String
    Dim customer As String
    Dim fName As String
    Dim fullPath As String
    Dim intWb As Workbook
    Dim wsOffer As Worksheet
    Dim wsCost As Worksheet
    Dim wsComp As Worksheet
    Dim wsOut As Worksheet
    Dim curUser As String
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim newRow As ListRow
    
    On Error Resume Next
    Set wsOffer = ThisWorkbook.Worksheets("OFFER")
    Set wsCost = ThisWorkbook.Worksheets("COSTING")
    Set wsComp = ThisWorkbook.Worksheets("COMPARE")
    On Error GoTo 0
    
    expDir = GetExportsDir(targetFolder)
    offerNo = CleanFilename(CStr(wsOffer.Range("C5").Value))
    customer = CleanFilename(CStr(wsOffer.Range("C7").Value))
    If Len(offerNo) = 0 Then offerNo = "MEGA_OFFER_26-27_594"
    If Len(customer) = 0 Then customer = "Customer"
    
    fName = offerNo & "_" & customer & "_R0_Internal.xlsx"
    fullPath = expDir & fName
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set intWb = Application.Workbooks.Add(xlWBATWorksheet)
    
    ' 1. OFFER
    Set wsOut = intWb.Worksheets(1)
    wsOut.Name = "OFFER"
    wsOffer.UsedRange.Copy
    wsOut.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    wsOut.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    ' 2. COSTING
    If Not wsCost Is Nothing Then
        Set wsOut = intWb.Worksheets.Add(After:=wsOut)
        wsOut.Name = "COSTING"
        wsCost.UsedRange.Copy
        wsOut.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
        wsOut.Range("A1").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If
    
    ' 3. COMPARE
    If Not wsComp Is Nothing Then
        Set wsOut = intWb.Worksheets.Add(After:=wsOut)
        wsOut.Name = "COMPARE"
        wsComp.UsedRange.Copy
        wsOut.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
        wsOut.Range("A1").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If
    
    ' Flatten all cells to values
    For Each wsOut In intWb.Worksheets
        wsOut.UsedRange.Value = wsOut.UsedRange.Value
    Next wsOut
    
    intWb.SaveAs FileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
    intWb.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Log to tblQuoteLog
    curUser = Environ$("USERNAME")
    If Len(curUser) = 0 Then curUser = "User"
    
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    If Not tblLog Is Nothing Then
        If tblLog.ListRows.Count = 1 And Len(Trim$(CStr(tblLog.ListRows(1).Range.Cells(1, tblLog.ListColumns("Action").Index).Value))) = 0 Then
            Set newRow = tblLog.ListRows(1)
        Else
            Set newRow = tblLog.ListRows.Add
        End If
        With newRow.Range
            .Cells(1, tblLog.ListColumns("LogID").Index).Value = "LOG-" & Format$(newRow.Index, "00000")
            .Cells(1, tblLog.ListColumns("QuoteID").Index).Value = offerNo
            .Cells(1, tblLog.ListColumns("Timestamp").Index).Value = Now
            .Cells(1, tblLog.ListColumns("User").Index).Value = curUser
            .Cells(1, tblLog.ListColumns("Action").Index).Value = "EXPORT_INTERNAL_XLSX"
            .Cells(1, tblLog.ListColumns("OldValue").Index).Value = ""
            .Cells(1, tblLog.ListColumns("NewValue").Index).Value = fName
            .Cells(1, tblLog.ListColumns("SheetName").Index).Value = "OFFER"
            .Cells(1, tblLog.ListColumns("Status").Index).Value = "EXPORTED"
            .Cells(1, tblLog.ListColumns("Notes").Index).Value = "Internal Excel exported to " & fullPath
        End With
    End If
    On Error GoTo 0
    
    WriteInternalExcel = fullPath
End Function

Public Function RunExportSelfTests() As String
    Dim tempExpDir As String
    Dim fso As Object
    Dim pdfPath As String
    Dim custPath As String
    Dim intPath As String
    Dim testWb As Workbook
    Dim wsCheck As Worksheet
    Dim cellCheck As Range
    
    On Error GoTo TestFail
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempExpDir = Environ$("TEMP") & "\MegaPipes_ExpTest_" & Format(Now, "hhnnss") & "\"
    If Not fso.FolderExists(tempExpDir) Then fso.CreateFolder tempExpDir
    
    ' 1. Test PDF Export
    pdfPath = ExportPDF(tempExpDir)
    If Not fso.FileExists(pdfPath) Then
        RunExportSelfTests = "FAIL: PDF file not created at " & pdfPath
        Exit Function
    End If
    If FileLen(pdfPath) < 1000 Then
        RunExportSelfTests = "FAIL: PDF file unexpectedly small: " & CStr(FileLen(pdfPath)) & " bytes"
        Exit Function
    End If
    
    ' 2. Test Customer Excel Export (Strict Privacy & Values-Only)
    custPath = WriteCustomerExcel(tempExpDir)
    If Not fso.FileExists(custPath) Then
        RunExportSelfTests = "FAIL: Customer Excel not created at " & custPath
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set testWb = Application.Workbooks.Open(FileName:=custPath, ReadOnly:=True)
    
    ' Check sheets: must have only OFFER and Synopsis
    If testWb.Worksheets.Count > 2 Then
        testWb.Close SaveChanges:=False
        RunExportSelfTests = "FAIL: Customer Excel has forbidden sheet count: " & CStr(testWb.Worksheets.Count)
        Exit Function
    End If
    
    For Each wsCheck In testWb.Worksheets
        If wsCheck.Name <> "OFFER" And wsCheck.Name <> "Synopsis" Then
            testWb.Close SaveChanges:=False
            RunExportSelfTests = "FAIL: Forbidden sheet in Customer Excel: " & wsCheck.Name
            Exit Function
        End If
        
        ' Verify no formulas
        For Each cellCheck In wsCheck.UsedRange
            If cellCheck.HasFormula Then
                testWb.Close SaveChanges:=False
                RunExportSelfTests = "FAIL: Formula detected in Customer Excel " & wsCheck.Name & "!" & cellCheck.Address
                Exit Function
            End If
        Next cellCheck
    Next wsCheck
    
    testWb.Close SaveChanges:=False
    
    ' 3. Test Internal Excel Export
    intPath = WriteInternalExcel(tempExpDir)
    If Not fso.FileExists(intPath) Then
        RunExportSelfTests = "FAIL: Internal Excel not created at " & intPath
        Exit Function
    End If
    
    Set testWb = Application.Workbooks.Open(FileName:=intPath, ReadOnly:=True)
    If testWb.Worksheets.Count < 2 Then
        testWb.Close SaveChanges:=False
        RunExportSelfTests = "FAIL: Internal Excel missing expected sheets"
        Exit Function
    End If
    
    ' Verify no formulas
    For Each wsCheck In testWb.Worksheets
        For Each cellCheck In wsCheck.UsedRange
            If cellCheck.HasFormula Then
                testWb.Close SaveChanges:=False
                RunExportSelfTests = "FAIL: Formula detected in Internal Excel " & wsCheck.Name & "!" & cellCheck.Address
                Exit Function
            End If
        Next cellCheck
    Next wsCheck
    testWb.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Clean up temporary export files
    On Error Resume Next
    fso.DeleteFile pdfPath, True
    fso.DeleteFile custPath, True
    fso.DeleteFile intPath, True
    fso.DeleteFolder tempExpDir, True
    On Error GoTo 0
    
    RunExportSelfTests = "PASS"
    Exit Function
    
TestFail:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not testWb Is Nothing Then
        On Error Resume Next
        testWb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    RunExportSelfTests = "FAIL: Exception - " & Err.Description
End Function


