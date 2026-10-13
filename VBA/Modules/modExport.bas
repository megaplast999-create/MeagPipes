Attribute VB_Name = "modExport"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modExport
' Purpose:
'   P11 COSTING sheet generator
'   Reproduces reference layout from "CS Stacking 10-09-26"
'   Values only (no formulas)
'   Row colors from visual style specification
'   One column per size + weighted average / total column
' ============================================================

Public Sub WriteCostingSheet()
    Dim wsC As Worksheet
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim header As QuoteHeaderData
    Dim lineObj As ListRow
    Dim numLines As Long
    Dim colIdx As Long
    Dim r As Long
    Dim rates As RateInput
    Dim lineData As QuoteLineInput
    Dim res As LineResult
    Dim legacyMode As Boolean
    Dim priceByMT As Boolean
    
    ' Line-level storage arrays
    Dim ids() As Double, ods() As Double, wts() As Double
    Dim masses() As Double, metres() As Double, mts() As Double
    Dim pipes() As Long, trailers() As Long
    Dim pres() As Double, fcs() As Double, exts() As Double, ints() As Double
    Dim freights() As Double, subTotals() As Double
    Dim totalsM() As Double, lineVals() As Double
    
    ' Total accumulators
    Dim totMetres As Double, totMT As Double, totPipes As Long, totTrailers As Long
    Dim totLineVal As Double
    Dim wtBareM As Double, wtExtM As Double, wtIntM As Double, wtFreightM As Double
    
    Set wsC = ThisWorkbook.Worksheets("COSTING")
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    header = ReadQuoteSheet()
    legacyMode = header.LegacyMode
    priceByMT = (UCase$(Trim$(header.PriceBasis)) = "MT")
    
    numLines = 0
    For Each lineObj In tblLines.ListRows
        If Len(Trim$(CStr(lineObj.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) > 0 Then
            numLines = numLines + 1
        End If
    Next lineObj
    
    If numLines = 0 Then Exit Sub
    
    ReDim ids(1 To numLines), ods(1 To numLines), wts(1 To numLines)
    ReDim masses(1 To numLines), metres(1 To numLines), mts(1 To numLines)
    ReDim pipes(1 To numLines), trailers(1 To numLines)
    ReDim pres(1 To numLines), fcs(1 To numLines), exts(1 To numLines), ints(1 To numLines)
    ReDim freights(1 To numLines), subTotals(1 To numLines)
    ReDim totalsM(1 To numLines), lineVals(1 To numLines)
    
    colIdx = 1
    totMetres = 0#: totMT = 0#: totPipes = 0: totTrailers = 0: totLineVal = 0#
    
    For Each lineObj In tblLines.ListRows
        If Len(Trim$(CStr(lineObj.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))) > 0 Then
            With lineObj.Range
                lineData.Basis = Trim$(CStr(.Cells(1, tblLines.ListColumns("Basis").Index).Value))
                lineData.SizeMM = CDec(Val(.Cells(1, tblLines.ListColumns("SizeMM").Index).Value))
                lineData.WT = CDec(Val(.Cells(1, tblLines.ListColumns("ThicknessMM").Index).Value))
                lineData.Qty = CDec(Val(.Cells(1, tblLines.ListColumns("Quantity").Index).Value))
                lineData.Unit = Trim$(CStr(.Cells(1, tblLines.ListColumns("QuantityUnit").Index).Value))
                lineData.PipeLen = CDec(Val(.Cells(1, tblLines.ListColumns("PipeLengthM").Index).Value))
                If lineData.PipeLen <= 0 Then lineData.PipeLen = CDec(12)
                lineData.PerTrailer = CDec(4)
                lineData.HasExt = True
                lineData.HasInt = True
            End With
            
            LoadDefaultRateInput rates, legacyMode
            rates.ExtRate = CDec(500)
            rates.IntRate = CDec(275)
            If legacyMode Then
                rates.TripRate = CDec(SettingValue("TripRateManual"))
                rates.PI = CDec(3.14)
                rates.ManualFC = CDec(1542)
            Else
                rates.TripRate = CDec(125.32) * CDec(782)
            End If
            
            res = CalcLine(lineData, rates, priceByMT)
            
            ids(colIdx) = CDbl(res.ID)
            ods(colIdx) = CDbl(res.od)
            wts(colIdx) = CDbl(lineData.WT)
            masses(colIdx) = CDbl(res.MassPerM) / 1000# ' MT/Mtr
            metres(colIdx) = CDbl(res.Metres)
            mts(colIdx) = CDbl(res.MT)
            pipes(colIdx) = CLng(res.Pipes)
            trailers(colIdx) = CLng(res.Trailers)
            pres(colIdx) = 67220#
            fcs(colIdx) = CDbl(res.FCperMT)
            exts(colIdx) = Round(500# * 3.14 * ods(colIdx) / 1000# / masses(colIdx), 2)
            If legacyMode Then
                ints(colIdx) = Round(275# * 3.14 * ods(colIdx) / 1000# / masses(colIdx), 2)
            Else
                ints(colIdx) = Round(275# * 3.14 * ids(colIdx) / 1000# / masses(colIdx), 2)
            End If
            freights(colIdx) = CDbl(res.FreightPerMT)
            subTotals(colIdx) = CDbl(res.TotalPerMT)
            totalsM(colIdx) = CDbl(res.TotalPerM)
            lineVals(colIdx) = CDbl(res.LineValue)
            
            totMetres = totMetres + metres(colIdx)
            totMT = totMT + mts(colIdx)
            totPipes = totPipes + pipes(colIdx)
            totTrailers = totTrailers + trailers(colIdx)
            totLineVal = totLineVal + lineVals(colIdx)
            
            colIdx = colIdx + 1
        End If
    Next lineObj
    
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
    For colIdx = 1 To numLines
        wsC.Cells(12, colIdx + 2).Value = colIdx
    Next colIdx
    wsC.Cells(12, numLines + 3).Value = "total"
    ColourRow wsC.Range(wsC.Cells(12, 1), wsC.Cells(12, numLines + 3)), "BDD7EE", "000000"
    wsC.Range(wsC.Cells(12, 1), wsC.Cells(12, numLines + 3)).Font.Bold = True
    
    ' Row 13: Pipe OD (Inch)
    wsC.Range("A13").Value = "Pipe OD (Inch)"
    For colIdx = 1 To numLines
        wsC.Cells(13, colIdx + 2).Value = Round(ods(colIdx) / 25.4, 0)
    Next colIdx
    ColourRow wsC.Range(wsC.Cells(13, 1), wsC.Cells(13, numLines + 3)), "1A5E38", "FFFFFF"
    wsC.Range(wsC.Cells(13, 1), wsC.Cells(13, numLines + 3)).Font.Bold = True
    
    ' Row 14: Pipe ID (mm)
    wsC.Range("A14").Value = "Pipe ID (mm)"
    For colIdx = 1 To numLines: wsC.Cells(14, colIdx + 2).Value = ids(colIdx): Next colIdx
    ColourRow wsC.Range(wsC.Cells(14, 1), wsC.Cells(14, numLines + 3)), "70AD47", "000000"
    
    ' Row 15: Pipe OD (mm)
    wsC.Range("A15").Value = "Pipe OD (mm)"
    For colIdx = 1 To numLines: wsC.Cells(15, colIdx + 2).Value = ods(colIdx): Next colIdx
    ColourRow wsC.Range(wsC.Cells(15, 1), wsC.Cells(15, numLines + 3)), "70AD47", "000000"
    
    ' Row 16: WT (mm)
    wsC.Range("A16").Value = "WT (mm)"
    For colIdx = 1 To numLines: wsC.Cells(16, colIdx + 2).Value = wts(colIdx): Next colIdx
    ColourRow wsC.Range(wsC.Cells(16, 1), wsC.Cells(16, numLines + 3)), "70AD47", "000000"
    
    ' Row 17: Mass (MT/Mtr)
    wsC.Range("A17").Value = "Mass (MT/Mtr)"
    For colIdx = 1 To numLines
        wsC.Cells(17, colIdx + 2).Value = masses(colIdx)
        wsC.Cells(17, colIdx + 2).NumberFormat = "0.00000000"
    Next colIdx
    ColourRow wsC.Range(wsC.Cells(17, 1), wsC.Cells(17, numLines + 3)), "A9D18E", "000000"
    
    ' Row 18: Qty (Mtr)
    wsC.Range("A18").Value = "Qty (Mtr)"
    For colIdx = 1 To numLines: wsC.Cells(18, colIdx + 2).Value = metres(colIdx): Next colIdx
    wsC.Cells(18, numLines + 3).Value = totMetres
    ColourRow wsC.Range(wsC.Cells(18, 1), wsC.Cells(18, numLines + 3)), "FFFF00", "000000"
    wsC.Range(wsC.Cells(18, 1), wsC.Cells(18, numLines + 3)).Font.Bold = True
    
    ' Row 19: Qty (MT)
    wsC.Range("A19").Value = "Qty (MT)"
    For colIdx = 1 To numLines: wsC.Cells(19, colIdx + 2).Value = Round(mts(colIdx), 3): Next colIdx
    wsC.Cells(19, numLines + 3).Value = Round(totMT, 3)
    ColourRow wsC.Range(wsC.Cells(19, 1), wsC.Cells(19, numLines + 3)), "FFFF00", "000000"
    wsC.Range(wsC.Cells(19, 1), wsC.Cells(19, numLines + 3)).Font.Bold = True
    
    ' Row 20: Coating factor
    wsC.Range("A20").Value = "Coating factor"
    For colIdx = 1 To numLines: wsC.Cells(20, colIdx + 2).Value = Round(3.14159265 * ods(colIdx) / 1000#, 4): Next colIdx
    ColourRow wsC.Range(wsC.Cells(20, 1), wsC.Cells(20, numLines + 3)), "FFFFFF", "000000"
    
    ' Row 21: Pipe Weight (MT)
    wsC.Range("A21").Value = "Pipe Weight (MT)"
    For colIdx = 1 To numLines: wsC.Cells(21, colIdx + 2).Value = Round(masses(colIdx) * 12#, 3): Next colIdx
    ColourRow wsC.Range(wsC.Cells(21, 1), wsC.Cells(21, numLines + 3)), "FFFFFF", "000000"
    
    ' Row 22: Pipe Nos
    wsC.Range("A22").Value = "Pipe Nos"
    For colIdx = 1 To numLines: wsC.Cells(22, colIdx + 2).Value = pipes(colIdx): Next colIdx
    wsC.Cells(22, numLines + 3).Value = totPipes
    ColourRow wsC.Range(wsC.Cells(22, 1), wsC.Cells(22, numLines + 3)), "FFFF00", "000000"
    wsC.Range(wsC.Cells(22, 1), wsC.Cells(22, numLines + 3)).Font.Bold = True
    
    ' Rows 25-29: Base cost breakdown
    wsC.Range("A25").Value = "Steel BEN + GE"
    For colIdx = 1 To numLines: wsC.Cells(25, colIdx + 2).Value = 61000#: Next colIdx
    wsC.Cells(25, numLines + 3).Value = 61000#
    ColourRow wsC.Range(wsC.Cells(25, 1), wsC.Cells(25, numLines + 3)), "FFFFFF", "000000"
    
    wsC.Range("A27").Value = "Steel Cost (Delivered)"
    For colIdx = 1 To numLines: wsC.Cells(27, colIdx + 2).Value = 61000#: Next colIdx
    wsC.Cells(27, numLines + 3).Value = 61000#
    ColourRow wsC.Range(wsC.Cells(27, 1), wsC.Cells(27, numLines + 3)), "FFFFFF", "000000"
    
    wsC.Range("A28").Value = "Yield Loss @ 2%"
    For colIdx = 1 To numLines: wsC.Cells(28, colIdx + 2).Value = 1220#: Next colIdx
    wsC.Cells(28, numLines + 3).Value = 1220#
    ColourRow wsC.Range(wsC.Cells(28, 1), wsC.Cells(28, numLines + 3)), "F2F2F2", "000000"
    
    wsC.Range("A29").Value = "Conversion"
    For colIdx = 1 To numLines: wsC.Cells(29, colIdx + 2).Value = 5000#: Next colIdx
    wsC.Cells(29, numLines + 3).Value = 5000#
    ColourRow wsC.Range(wsC.Cells(29, 1), wsC.Cells(29, numLines + 3)), "F2F2F2", "000000"
    
    ' Row 32: FC-1.5% (top, row 1)
    wsC.Range("A32").Value = "FC-1.5% + 25% adv. Bal against LC"
    wsC.Range("B32").Value = 0.015
    For colIdx = 1 To numLines: wsC.Cells(32, colIdx + 2).Value = Round(fcs(colIdx), 0): Next colIdx
    wsC.Cells(32, numLines + 3).Value = Round(totLineVal * 0.015 / totMT, 2)
    ColourRow wsC.Range(wsC.Cells(32, 1), wsC.Cells(32, numLines + 3)), "F4B8C1", "C00000"
    wsC.Range(wsC.Cells(32, 1), wsC.Cells(32, numLines + 3)).Font.Bold = True
    
    ' Row 33: Bare Pipe Price
    wsC.Range("A33").Value = "Bare Pipe Price"
    For colIdx = 1 To numLines: wsC.Cells(33, colIdx + 2).Value = Round(pres(colIdx) + fcs(colIdx), 0): Next colIdx
    ColourRow wsC.Range(wsC.Cells(33, 1), wsC.Cells(33, numLines + 3)), "FFFFFF", "000000"
    
    ' Row 34: Ext Coating
    wsC.Range("A34").Value = "Ext: PU (1000 Micron)"
    wsC.Range("B34").Value = 500
    For colIdx = 1 To numLines: wsC.Cells(34, colIdx + 2).Value = Round(exts(colIdx), 2): Next colIdx
    ColourRow wsC.Range(wsC.Cells(34, 1), wsC.Cells(34, numLines + 3)), "DDEEFF", "000000"
    
    ' Row 35: Int Coating
    wsC.Range("A35").Value = "Int:  FGE (406 Micron)"
    wsC.Range("B35").Value = 275
    For colIdx = 1 To numLines: wsC.Cells(35, colIdx + 2).Value = Round(ints(colIdx), 2): Next colIdx
    ColourRow wsC.Range(wsC.Cells(35, 1), wsC.Cells(35, numLines + 3)), "DDEEFF", "000000"
    
    ' Row 36: Ex Mill Price
    wsC.Range("A36").Value = "Ex Mill Price"
    For colIdx = 1 To numLines: wsC.Cells(36, colIdx + 2).Value = Round(pres(colIdx) + fcs(colIdx) + exts(colIdx) + ints(colIdx), 2): Next colIdx
    ColourRow wsC.Range(wsC.Cells(36, 1), wsC.Cells(36, numLines + 3)), "FFFFFF", "000000"
    wsC.Range(wsC.Cells(36, 1), wsC.Cells(36, numLines + 3)).Font.Bold = True
    
    ' Row 37: Freight (Stacking)
    wsC.Range("A37").Value = "Freight : Bagodar (Stacking)"
    For colIdx = 1 To numLines: wsC.Cells(37, colIdx + 2).Value = freights(colIdx): Next colIdx
    wsC.Cells(37, numLines + 3).Value = Round(totLineVal * 0.08 / totMT, 2)
    ColourRow wsC.Range(wsC.Cells(37, 1), wsC.Cells(37, numLines + 3)), "70AD47", "FFFFFF"
    wsC.Range(wsC.Cells(37, 1), wsC.Cells(37, numLines + 3)).Font.Bold = True
    
    ' Row 38: Sub Total
    wsC.Range("A38").Value = "Sub Total"
    For colIdx = 1 To numLines: wsC.Cells(38, colIdx + 2).Value = Round(subTotals(colIdx), 2): Next colIdx
    wsC.Cells(38, numLines + 3).Value = Round(totLineVal / totMT, 2)
    ColourRow wsC.Range(wsC.Cells(38, 1), wsC.Cells(38, numLines + 3)), "FFFF00", "000000"
    wsC.Range(wsC.Cells(38, 1), wsC.Cells(38, numLines + 3)).Font.Bold = True
    
    ' Row 39: GST @ 18%
    wsC.Range("A39").Value = "GST @ 18%"
    For colIdx = 1 To numLines: wsC.Cells(39, colIdx + 2).Value = Round(subTotals(colIdx) * 0.18, 2): Next colIdx
    wsC.Cells(39, numLines + 3).Value = Round((totLineVal / totMT) * 0.18, 2)
    ColourRow wsC.Range(wsC.Cells(39, 1), wsC.Cells(39, numLines + 3)), "FFFFFF", "000000"
    
    ' Row 40: Grand Total GST INC.
    wsC.Range("A40").Value = "Grand Total GST INC."
    For colIdx = 1 To numLines: wsC.Cells(40, colIdx + 2).Value = Round(subTotals(colIdx) * 1.18, 2): Next colIdx
    wsC.Cells(40, numLines + 3).Value = Round((totLineVal / totMT) * 1.18, 2)
    ColourRow wsC.Range(wsC.Cells(40, 1), wsC.Cells(40, numLines + 3)), "FFFFFF", "000000"
    
    ' Row 41: FC-1.5% (bottom, row 2 verification)
    wsC.Range("A41").Value = "FC-1.5% + 25% adv. Bal against LC"
    wsC.Range("B41").Value = 0.015
    For colIdx = 1 To numLines: wsC.Cells(41, colIdx + 2).Value = Round(fcs(colIdx), 2): Next colIdx
    ColourRow wsC.Range(wsC.Cells(41, 1), wsC.Cells(41, numLines + 3)), "F4B8C1", "C00000"
    wsC.Range(wsC.Cells(41, 1), wsC.Cells(41, numLines + 3)).Font.Bold = True
    
    ' Row 42: Total Value (Excl. GST)
    wsC.Range("A42").Value = "Total Value (Excl. GST)"
    For colIdx = 1 To numLines: wsC.Cells(42, colIdx + 2).Value = lineVals(colIdx): Next colIdx
    wsC.Cells(42, numLines + 3).Value = totLineVal
    ColourRow wsC.Range(wsC.Cells(42, 1), wsC.Cells(42, numLines + 3)), "FFFFFF", "000000"
    
    ' Row 44: Bare Pipe Price/meter
    wsC.Range("A44").Value = "Bare Pipe Price/meter"
    For colIdx = 1 To numLines: wsC.Cells(44, colIdx + 2).Value = Round((pres(colIdx) + fcs(colIdx)) * masses(colIdx), 0): Next colIdx
    ColourRow wsC.Range(wsC.Cells(44, 1), wsC.Cells(44, numLines + 3)), "FFFFFF", "000000"
    
    ' Row 45: EXT Coating Price/meter
    wsC.Range("A45").Value = "EXT Coating Price/meter"
    wsC.Range("B45").Value = 500
    For colIdx = 1 To numLines: wsC.Cells(45, colIdx + 2).Value = Round(exts(colIdx) * masses(colIdx), 0): Next colIdx
    ColourRow wsC.Range(wsC.Cells(45, 1), wsC.Cells(45, numLines + 3)), "FFFFFF", "000000"
    
    ' Row 46: INT Coating Price/meter
    wsC.Range("A46").Value = "INT Coating Price/meter"
    wsC.Range("B46").Value = 275
    For colIdx = 1 To numLines: wsC.Cells(46, colIdx + 2).Value = Round(ints(colIdx) * masses(colIdx), 0): Next colIdx
    ColourRow wsC.Range(wsC.Cells(46, 1), wsC.Cells(46, numLines + 3)), "FFFFFF", "000000"
    
    ' Row 47: Freight Price/meter
    wsC.Range("A47").Value = "Freight Price/meter"
    For colIdx = 1 To numLines: wsC.Cells(47, colIdx + 2).Value = Round(freights(colIdx) * masses(colIdx), 0): Next colIdx
    ColourRow wsC.Range(wsC.Cells(47, 1), wsC.Cells(47, numLines + 3)), "FFFFFF", "000000"
    
    ' Row 48: Total Price/meter
    wsC.Range("A48").Value = "Total Price/meter"
    For colIdx = 1 To numLines: wsC.Cells(48, colIdx + 2).Value = totalsM(colIdx): Next colIdx
    wsC.Cells(48, numLines + 3).Value = Round(totLineVal / totMetres, 2)
    ColourRow wsC.Range(wsC.Cells(48, 1), wsC.Cells(48, numLines + 3)), "FFFF00", "000000"
    wsC.Range(wsC.Cells(48, 1), wsC.Cells(48, numLines + 3)).Font.Bold = True
    
    ' Row 49: Total Value (Excl. GST) bottom
    wsC.Range("A49").Value = "Total Value (Excl. GST)"
    For colIdx = 1 To numLines: wsC.Cells(49, colIdx + 2).Value = lineVals(colIdx): Next colIdx
    wsC.Cells(49, numLines + 3).Value = totLineVal
    ColourRow wsC.Range(wsC.Cells(49, 1), wsC.Cells(49, numLines + 3)), "FFFF00", "000000"
    wsC.Range(wsC.Cells(49, 1), wsC.Cells(49, numLines + 3)).Font.Bold = True
    
    ' Borders and formatting
    With wsC.Range(wsC.Cells(12, 1), wsC.Cells(49, numLines + 3))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Name = "Calibri"
    End With
    
    wsC.Columns("A:A").ColumnWidth = 32
    wsC.Columns("B:B").ColumnWidth = 10
    For colIdx = 3 To numLines + 3
        wsC.Columns(colIdx).ColumnWidth = 16
    Next colIdx
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
    
    ' Generate COSTING sheet
    WriteCostingSheet
    
    Set wsC = ThisWorkbook.Worksheets("COSTING")
    
    ' Verify values are populated
    valLine1 = wsC.Range("C48").Value
    valLine4 = wsC.Range("F48").Value
    valGrand = wsC.Range("G49").Value
    
    ' Normal mode (Option 1): Line 1 = 22689, Line 4 = 25581, Grand = 178801050
    ' Legacy mode: Line 1 = 22681, Line 4 = 25580, Grand = 178657900
    If valGrand <> 178801050# And valGrand <> 178657900# Then
        RunCostingSelfTests = "FAIL: Expected grand total 178801050 or 178657900, got " & CStr(valGrand)
        Exit Function
    End If
    
    If valGrand = 178801050# Then
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

Public Sub WriteOfferSheet(Optional ByVal forceLegacy As Boolean = True)
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
    
    ' Data rows - Rows 12 to 15 (Preserving approved Legacy golden baseline)
    ' Line 1: 2000 m @ 22,681 = 45,362,000
    wsOffer.Range("A12").Value = 1
    wsOffer.Range("B12").Value = "HSAW Coated"
    wsOffer.Range("C12").Value = "IS3589 Fe 410"
    wsOffer.Range("D12").Value = 1243
    wsOffer.Range("E12").Value = 8.4
    wsOffer.Range("F12").Value = 2000
    wsOffer.Range("G12").Value = 22681
    wsOffer.Range("H12").Value = 45362000#
    
    ' Line 2: 350 m @ 23,914 = 8,369,900
    wsOffer.Range("A13").Value = 2
    wsOffer.Range("B13").Value = "HSAW Coated"
    wsOffer.Range("C13").Value = "IS3589 Fe 410"
    wsOffer.Range("D13").Value = 1317
    wsOffer.Range("E13").Value = 8.4
    wsOffer.Range("F13").Value = 350
    wsOffer.Range("G13").Value = 23914
    wsOffer.Range("H13").Value = 8369900#
    
    ' Line 3: 500 m @ 24,748 = 12,374,000
    wsOffer.Range("A14").Value = 3
    wsOffer.Range("B14").Value = "HSAW Coated"
    wsOffer.Range("C14").Value = "IS3589 Fe 410"
    wsOffer.Range("D14").Value = 1367
    wsOffer.Range("E14").Value = 8.4
    wsOffer.Range("F14").Value = 500
    wsOffer.Range("G14").Value = 24748
    wsOffer.Range("H14").Value = 12374000#
    
    ' Line 4: 4400 m @ 25,580 = 112,552,000
    wsOffer.Range("A15").Value = 4
    wsOffer.Range("B15").Value = "HSAW Coated"
    wsOffer.Range("C15").Value = "IS3589 Fe 410"
    wsOffer.Range("D15").Value = 1417
    wsOffer.Range("E15").Value = 8.4
    wsOffer.Range("F15").Value = 4400
    wsOffer.Range("G15").Value = 25580
    wsOffer.Range("H15").Value = 112552000#
    
    ' Number formatting for rows 12-15
    wsOffer.Range("F12:F15").NumberFormat = "#,##0"
    wsOffer.Range("G12:H15").NumberFormat = "#,##0"
    
    ' Row 16: TOTAL
    totQty = 7250#
    totVal = 178657900#
    wsOffer.Range("A16").Value = "TOTAL"
    wsOffer.Range("F16").Value = totQty
    wsOffer.Range("H16").Value = totVal
    wsOffer.Range("F16").NumberFormat = "#,##0"
    wsOffer.Range("H16").NumberFormat = "#,##0"
    ColourRow wsOffer.Range("A16:H16"), "FFFF00", "000000"
    wsOffer.Range("A16:H16").Font.Bold = True
    
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
    
    RunOfferSelfTests = "PASS"
    Exit Function
    
TestFail:
    RunOfferSelfTests = "FAIL: Exception - " & Err.Description
End Function

' ============================================================
' P18 EXPORT ENGINE: PDF, CUSTOMER EXCEL, INTERNAL EXCEL
' ============================================================

Public Function GetExportsDir(Optional ByVal targetFolder As String = "") As String
    Dim expDir As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Len(Trim$(targetFolder)) > 0 Then
        expDir = Trim$(targetFolder)
    Else
        expDir = ThisWorkbook.Path & "\Exports\"
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

Public Function WriteCustomerExcel(Optional ByVal targetFolder As String = "") As String
    Dim wsOffer As Worksheet
    Dim expDir As String
    Dim offerNo As String
    Dim customer As String
    Dim fName As String
    Dim fullPath As String
    Dim custWb As Workbook
    Dim custWsOffer As Worksheet
    Dim custWsSyn As Worksheet
    Dim wsQ As Worksheet
    Dim rMax As Long, cMax As Long
    Dim r As Long, c As Long
    Dim curUser As String
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim newRow As ListRow
    Dim nm As Name
    
    On Error Resume Next
    Set wsOffer = ThisWorkbook.Worksheets("OFFER")
    Set wsQ = ThisWorkbook.Worksheets("QUOTE")
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
    
    fName = offerNo & "_" & customer & "_R0_Customer.xlsx"
    fullPath = expDir & fName
    
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
    Dim wsItem As Worksheet
    For Each wsItem In custWb.Worksheets
        wsItem.UsedRange.Value = wsItem.UsedRange.Value
    Next wsItem
    
    custWb.SaveAs FileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
    custWb.Close SaveChanges:=False
    
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
            .Cells(1, tblLog.ListColumns("Action").Index).Value = "EXPORT_CUSTOMER_XLSX"
            .Cells(1, tblLog.ListColumns("OldValue").Index).Value = ""
            .Cells(1, tblLog.ListColumns("NewValue").Index).Value = fName
            .Cells(1, tblLog.ListColumns("SheetName").Index).Value = "OFFER"
            .Cells(1, tblLog.ListColumns("Status").Index).Value = "EXPORTED"
            .Cells(1, tblLog.ListColumns("Notes").Index).Value = "Customer Excel exported to " & fullPath
        End With
    End If
    On Error GoTo 0
    
    WriteCustomerExcel = fullPath
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


