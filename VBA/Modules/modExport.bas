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
