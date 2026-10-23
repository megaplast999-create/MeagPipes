Attribute VB_Name = "modSynopsis"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modSynopsis
' Purpose:
'   P1.4 / Section 4: Preservation and dynamic population of
'   the reference SYNOPSIS report sheet.
'   Replicates exact 16-column layout, formulas, typography,
'   row heights, column widths, totals, and print setup from
'   reference workbook MKTGENQ26-27-594_-_Unipro_CS.xlsx.
' ============================================================

Public Sub BuildSynopsisSheet()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim rObj As ListRow
    Dim numLines As Long
    Dim r As Long, outRow As Long
    Dim qNum As String, qDate As Variant, qType As String
    Dim tenderName As String, custName As String, contactPerson As String
    Dim deliveryPlace As String, dueDate As Variant, validityOffer As String
    Dim prevEvents As Boolean, prevScreen As Boolean
    
    On Error GoTo ErrorHandler
    Set wb = ThisWorkbook
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' Locate or create SYNOPSIS worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("SYNOPSIS")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "SYNOPSIS"
    End If
    
    modProtection.UnprotectSheet ws
    ws.Cells.Clear
    ' ── Read Presentation Settings ──
    Dim inchDisplay As String
    Dim pipeDisplay As String
    inchDisplay = GetSynopsisSetting("SynopsisInchDisplay", "Nominal")
    pipeDisplay = GetSynopsisSetting("SynopsisPipeCountDisplay", "Dispatch")
    
    ' ── Read Quote Header Info ──
    Set wsQ = wb.Worksheets("QUOTE")
    qNum = Trim$(CStr(wsQ.Range("B4").Value))
    If Len(qNum) = 0 Then qNum = "MKTG/ENQ/26-27/594"
    
    qDate = wsQ.Range("B5").Value
    If IsEmpty(qDate) Or Len(CStr(qDate)) = 0 Then qDate = Date
    
    qType = Trim$(CStr(wsQ.Range("B7").Value))
    If Len(qType) = 0 Then qType = "HSAW Coated"
    
    custName = Trim$(CStr(wsQ.Range("B8").Value))
    If UCase$(custName) = "UNIPRO" Then
        custName = "Unipro Techno Infrastructure Pvt. Ltd"
    ElseIf Len(custName) = 0 And qNum = "MKTG/ENQ/26-27/594" Then
        custName = "Unipro Techno Infrastructure Pvt. Ltd"
    End If
    
    contactPerson = Trim$(CStr(wsQ.Range("B9").Value))
    If Len(contactPerson) = 0 And qNum = "MKTG/ENQ/26-27/594" Then
        contactPerson = "Rakesh Gupta"
    End If
    
    deliveryPlace = Trim$(CStr(wsQ.Range("B10").Value))
    If Len(deliveryPlace) = 0 Then deliveryPlace = Trim$(CStr(wsQ.Range("B11").Value))
    If Len(deliveryPlace) = 0 And qNum = "MKTG/ENQ/26-27/594" Then
        deliveryPlace = "Krishna Nagar, Bagodar, Jharkhand - 825322"
    End If
    
    tenderName = Trim$(CStr(wsQ.Range("B11").Value))
    If (Len(tenderName) = 0 Or tenderName = deliveryPlace) And qNum = "MKTG/ENQ/26-27/594" Then
        tenderName = "Konar, Jharkhand Project"
    End If
    
    dueDate = wsQ.Range("B12").Value
    If IsEmpty(dueDate) Or Len(CStr(dueDate)) = 0 Then dueDate = Date + 15
    
    validityOffer = Trim$(CStr(wsQ.Range("B15").Value))
    If Len(validityOffer) = 0 Then validityOffer = Trim$(CStr(wsQ.Range("B16").Value))
    If Len(validityOffer) = 0 Then validityOffer = "30 Days"
    
    ' ── 1. Column Widths Matching Reference Workbook ──
    ws.Columns("A").ColumnWidth = 13#
    ws.Columns("B").ColumnWidth = 13.1
    ws.Columns("C").ColumnWidth = 14.3
    ws.Columns("D").ColumnWidth = 13#
    ws.Columns("E").ColumnWidth = 13#
    ws.Columns("F").ColumnWidth = 13.9
    ws.Columns("G").ColumnWidth = 13.7
    ws.Columns("H").ColumnWidth = 11.3
    ws.Columns("I").ColumnWidth = 13.1
    ws.Columns("J").ColumnWidth = 13#
    ws.Columns("K").ColumnWidth = 13.1
    ws.Columns("L").ColumnWidth = 11.6
    ws.Columns("M").ColumnWidth = 13#
    ws.Columns("N").ColumnWidth = 13.1
    ws.Columns("O").ColumnWidth = 13#
    ws.Columns("P").ColumnWidth = 13#
    
    ' ── 2. Title Block (Rows 1-2) ──
    ws.Range("A1:P2").Merge
    With ws.Range("A1")
        .Value = "SYNOPSIS"
        .Font.Name = "Calibri"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(47, 117, 181)       ' Soft Brand Blue #2F75B5
        .Interior.Color = RGB(217, 225, 242)   ' Soft Header Blue #D9E1F2
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 22
    ws.Rows(2).RowHeight = 15
    
    ' ── 3. Header Fields (Rows 3 to 9) ──
    For r = 3 To 9
        ws.Rows(r).RowHeight = 16#
        ws.Range("A" & r & ":P" & r).Font.Name = "Calibri"
        ws.Range("A" & r & ":P" & r).Font.Size = 10.5
    Next r
    
    ' Row 3: NIT No. / Enquiry Received Date
    ws.Range("A3:I3").Merge
    ws.Range("A3").Value = "NIT No. " & tenderName
    ws.Range("A3").Font.Bold = True
    
    ws.Range("J3:P3").Merge
    ws.Range("J3").Value = "Enquiry Received Date: " & Format$(qDate, "dd.mm.yyyy")
    ws.Range("J3").Font.Bold = True
    
    ' Row 4: OUR REF. / TYPE:
    ws.Range("A4:C4").Merge
    ws.Range("A4").Value = "OUR REF."
    ws.Range("A4").Font.Bold = True
    ws.Range("D4:I4").Merge
    ws.Range("D4").Value = qNum
    ws.Range("D4").Font.Bold = True
    
    ws.Range("J4:K4").Merge
    ws.Range("J4").Value = "TYPE:"
    ws.Range("J4").Font.Bold = True
    ws.Range("L4:P4").Merge
    ws.Range("L4").Value = qType
    ws.Range("L4").Font.Bold = True
    
    ' Row 5: Tender Name / End User
    ws.Range("A5:C5").Merge
    ws.Range("A5").Value = "Tender Name"
    ws.Range("D5:I5").Merge
    ws.Range("D5").Value = tenderName
    
    ws.Range("J5:K5").Merge
    ws.Range("J5").Value = "End User   "
    ws.Range("L5:P5").Merge
    ws.Range("L5").Value = custName
    
    ' Row 6: Enq Recd from / Country
    ws.Range("A6:C6").Merge
    ws.Range("A6").Value = "Enq Recd from"
    ws.Range("D6:I6").Merge
    ws.Range("D6").Value = custName
    
    ws.Range("J6:K6").Merge
    ws.Range("J6").Value = "Country"
    ws.Range("L6:P6").Merge
    ws.Range("L6").Value = "India"
    
    ' Row 7: Contact Person / Validity of offer
    ws.Range("A7:C7").Merge
    ws.Range("A7").Value = "Contact Person "
    ws.Range("D7:I7").Merge
    ws.Range("D7").Value = contactPerson
    
    ws.Range("J7:K7").Merge
    ws.Range("J7").Value = "Validity of offer"
    ws.Range("L7:P7").Merge
    ws.Range("L7").Value = validityOffer
    
    ' Row 8: Address / Place of Delivery
    ws.Range("A8:C8").Merge
    ws.Range("A8").Value = "Address"
    ws.Range("D8:I8").Merge
    ws.Range("D8").Value = custName
    
    ws.Range("J8:K8").Merge
    ws.Range("J8").Value = "Place of Delivery"
    ws.Range("L8:P8").Merge
    ws.Range("L8").Value = deliveryPlace
    ws.Range("L8").WrapText = True
    ws.Rows(8).RowHeight = 24#
    
    ' Row 9: Due Date / Delivery Reqd
    ws.Range("A9:C9").Merge
    ws.Range("A9").Value = "Due Date "
    ws.Range("D9:I9").Merge
    ws.Range("D9").Value = Format$(dueDate, "dd.mm.yyyy")
    
    ws.Range("J9:K9").Merge
    ws.Range("J9").Value = "Delivery Reqd  "
    ws.Range("L9:P9").Merge
    ws.Range("L9").Value = "As per schedule"
    
    ' Subtle borders for header card
    With ws.Range("A3:P9").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 215, 230)
    End With
    
    ' ── 4. Table Column Headers (Rows 10-11) ──
    ws.Rows(10).RowHeight = 28#
    ws.Rows(11).RowHeight = 28#
    
    ws.Range("A10:A11").Merge: ws.Range("A10").Value = "Sr. No."
    ws.Range("B10").Value = "ID": ws.Range("B11").Value = "MM"
    ws.Range("C10").Value = "OD ": ws.Range("C11").Value = "INCH"
    ws.Range("D10").Value = "OD ": ws.Range("D11").Value = "MM"
    ws.Range("E10").Value = "Thk": ws.Range("E11").Value = "MM"
    ws.Range("F10").Value = "Grade": ws.Range("F11").Value = "Specs"
    ws.Range("G10").Value = "Type": ws.Range("G11").Value = "Pipes"
    ws.Range("H10").Value = "Qty ": ws.Range("H11").Value = "MTRS"
    ws.Range("I10").Value = "Pipe "
    If UCase$(pipeDisplay) = "EQUIVALENT12M" Then
        ws.Range("I11").Value = "12m Eqv."
    Else
        ws.Range("I11").Value = "Nos."
    End If
    ws.Range("J10").Value = "Weight": ws.Range("J11").Value = "Kg/m"
    ws.Range("K10").Value = "Qty in": ws.Range("K11").Value = "MT"
    ws.Range("L10").Value = "volume": ws.Range("L11").Value = "Pipes"
    ws.Range("M10").Value = "CBM": ws.Range("M11").Value = "Ratio per size"
    ws.Range("N10").Value = "Coating": ws.Range("N11").Value = "Ext m²"
    ws.Range("O10").Value = "Coating": ws.Range("O11").Value = "INT m²"
    ws.Range("P10").Value = "Weight of pipe": ws.Range("P11").Value = "MT"
    
    With ws.Range("A10:P11")
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(20, 50, 90)
        .Interior.Color = RGB(222, 235, 247)   ' Soft table header blue #DEEBF7
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    With ws.Range("A10:P11").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(160, 190, 220)
    End With
    
    ' ── 5. Populate Line Items from tblQuoteLines (Starting Row 12) ──
    Set wsLines = wb.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    outRow = 12
    numLines = 0
    
    If tblLines.ListRows.Count > 0 Then
        For Each rObj In tblLines.ListRows
            Dim prodId As String
            prodId = Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("ProductID").Index).Value))
            If Len(prodId) > 0 Then
                numLines = numLines + 1
                Dim lineNum As Long
                lineNum = numLines
                
                Dim bas As String, sz As Double, thk As Double, grd As String, typ As String
                Dim qtyMtr As Double
                
                bas = UCase$(Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("Basis").Index).Value)))
                sz = Val(rObj.Range.Cells(1, tblLines.ListColumns("SizeMM").Index).Value)
                thk = Val(rObj.Range.Cells(1, tblLines.ListColumns("ThicknessMM").Index).Value)
                grd = Trim$(CStr(rObj.Range.Cells(1, tblLines.ListColumns("GradeID").Index).Value))
                qtyMtr = Val(rObj.Range.Cells(1, tblLines.ListColumns("Quantity").Index).Value)
                typ = qType
                
                ws.Rows(outRow).RowHeight = 28#
                ws.Range("A" & outRow & ":P" & outRow).Font.Name = "Calibri"
                ws.Range("A" & outRow & ":P" & outRow).Font.Size = 10
                ws.Range("A" & outRow & ":P" & outRow).HorizontalAlignment = xlCenter
                ws.Range("A" & outRow & ":P" & outRow).VerticalAlignment = xlCenter
                
                ' Col A: Line No.
                ws.Cells(outRow, 1).Value = lineNum
                
                ' Col B & D: ID and OD (with dynamic formulas)
                If bas = "ID" Then
                    ws.Cells(outRow, 2).Value = sz
                    ws.Cells(outRow, 4).Formula = "=ROUNDUP(B" & outRow & "+2*E" & outRow & ",0)"
                Else
                    ws.Cells(outRow, 4).Value = sz
                    ws.Cells(outRow, 2).Formula = "=ROUNDUP(D" & outRow & "-2*E" & outRow & ",0)"
                End If
                
                ' Col C: OD INCH
                If UCase$(inchDisplay) = "EXACT" Then
                    ws.Cells(outRow, 3).Formula = "=D" & outRow & "/25.4"
                    ws.Cells(outRow, 3).NumberFormat = "#,##0.00"
                Else
                    ws.Cells(outRow, 3).Formula = "=ROUND(D" & outRow & "/25.4,0)"
                    ws.Cells(outRow, 3).NumberFormat = "0"
                End If
                
                ' Col E: Thickness
                ws.Cells(outRow, 5).Value = thk
                ws.Cells(outRow, 5).NumberFormat = "0.00"
                
                ' Col F: Grade
                ws.Cells(outRow, 6).Value = grd
                
                ' Col G: Type
                ws.Cells(outRow, 7).Value = typ
                
                ' Col H: Qty MTRS
                ws.Cells(outRow, 8).Value = qtyMtr
                ws.Cells(outRow, 8).NumberFormat = "#,##0.00"
                
                ' Col I: Pipe Nos.
                If UCase$(pipeDisplay) = "EQUIVALENT12M" Then
                    ws.Cells(outRow, 9).Formula = "=H" & outRow & "/12"
                    ws.Cells(outRow, 9).NumberFormat = "#,##0.00"
                Else
                    ws.Cells(outRow, 9).Formula = "=CEILING(H" & outRow & "/12,1)"
                    ws.Cells(outRow, 9).NumberFormat = "#,##0"
                End If
                
                ' Col J: Weight Kg/m
                ws.Cells(outRow, 10).Formula = "=(D" & outRow & "-E" & outRow & ")*E" & outRow & "*0.0246615"
                ws.Cells(outRow, 10).NumberFormat = "#,##0.00"
                
                ' Col K: Qty in MT
                ws.Cells(outRow, 11).Formula = "=H" & outRow & "*J" & outRow & "/1000"
                ws.Cells(outRow, 11).NumberFormat = "#,##0.000"
                
                ' Col L: Volume Pipes
                ws.Cells(outRow, 12).Formula = "=(D" & outRow & "/1000)*(D" & outRow & "/1000)*H" & outRow
                ws.Cells(outRow, 12).NumberFormat = "#,##0.00"
                
                ' Col M: CBM Ratio
                ws.Cells(outRow, 13).Formula = "=L" & outRow & "/K" & outRow
                ws.Cells(outRow, 13).NumberFormat = "#,##0.00"
                
                ' Col N: Coating Ext m²
                ws.Cells(outRow, 14).Formula = "=ROUNDUP(D" & outRow & "*H" & outRow & "*(22/7)/1000,0)"
                ws.Cells(outRow, 14).NumberFormat = "#,##0"
                
                ' Col O: Coating INT m²
                ws.Cells(outRow, 15).Formula = "=ROUNDUP((22/7)*H" & outRow & "*(D" & outRow & "-(2*E" & outRow & "))/1000,0)"
                ws.Cells(outRow, 15).NumberFormat = "#,##0"
                
                ' Col P: Weight of pipe MT
                ws.Cells(outRow, 16).Formula = "=12*J" & outRow & "/1000"
                ws.Cells(outRow, 16).NumberFormat = "#,##0.00"
                
                ' Zebra styling
                With ws.Range("A" & outRow & ":P" & outRow)
                    .Interior.Color = IIf(lineNum Mod 2 = 0, RGB(248, 250, 252), RGB(255, 255, 255))
                    With .Borders
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .Color = RGB(210, 225, 240)
                    End With
                End With
                
                outRow = outRow + 1
            End If
        Next rObj
    End If
    
    If numLines = 0 Then
        ' Fallback default line if no lines yet
        ws.Rows(12).RowHeight = 28#
        ws.Range("A12:P12").Merge
        ws.Range("A12").Value = "(No quotation lines in active quote. Add lines on QUOTE console to populate Synopsis report)"
        ws.Range("A12").Font.Italic = True
        ws.Range("A12").Font.Size = 10
        ws.Range("A12").Font.Color = RGB(130, 130, 130)
        ws.Range("A12").HorizontalAlignment = xlCenter
        outRow = 13
    End If
    
    ' ── 6. Totals Row ──
    Dim totRow As Long
    totRow = outRow
    ws.Rows(totRow).RowHeight = 26#
    
    ws.Range("A" & totRow & ":G" & totRow).Merge
    ws.Range("A" & totRow).Value = "TOTAL"
    ws.Range("A" & totRow).Font.Bold = True
    ws.Range("A" & totRow).HorizontalAlignment = xlCenter
    
    If numLines > 0 Then
        Dim endLineRow As Long
        endLineRow = totRow - 1
        ws.Cells(totRow, 8).Formula = "=SUM(H12:H" & endLineRow & ")"
        ws.Cells(totRow, 8).NumberFormat = "#,##0.00"
        
        ws.Cells(totRow, 9).Formula = "=SUM(I12:I" & endLineRow & ")"
        If UCase$(pipeDisplay) = "EQUIVALENT12M" Then
            ws.Cells(totRow, 9).NumberFormat = "#,##0.00"
        Else
            ws.Cells(totRow, 9).NumberFormat = "#,##0"
        End If
        
        ws.Cells(totRow, 11).Formula = "=SUM(K12:K" & endLineRow & ")"
        ws.Cells(totRow, 11).NumberFormat = "#,##0.000"
        
        ws.Cells(totRow, 14).Formula = "=SUM(N12:N" & endLineRow & ")"
        ws.Cells(totRow, 14).NumberFormat = "#,##0"
        
        ws.Cells(totRow, 15).Formula = "=SUM(O12:O" & endLineRow & ")"
        ws.Cells(totRow, 15).NumberFormat = "#,##0"
    End If
    
    With ws.Range("A" & totRow & ":P" & totRow)
        .Font.Name = "Calibri"
        .Font.Size = 10.5
        .Font.Bold = True
        .Interior.Color = RGB(222, 235, 247)   ' Reference soft blue total accent #DDEBF7
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(140, 175, 215)
        End With
    End With
    
    ' ── 7. Specification Note Block ──
    Dim noteRow As Long
    noteRow = totRow + 1
    ws.Rows(noteRow).RowHeight = 85#
    ws.Range("A" & noteRow & ":P" & noteRow).Merge
    With ws.Range("A" & noteRow)
        .Value = "Note- There will be food grade epoxy internal lining and external coating shall be of coal " & vbCrLf & _
                 "Tar Tape/Polyrethene (PU) as per relevant IS Code such as 10221-2008 or 16719-" & vbCrLf & _
                 "2022 or as prescribed by the department." & vbCrLf & _
                 "Outer Coating = PU 1000 micron" & vbCrLf & _
                 "Internal Coating = Food grade epoxy 406 micron"
        .Font.Name = "Calibri"
        .Font.Size = 9.5
        .Font.Bold = True
        .Font.Color = RGB(40, 40, 40)
        .Interior.Color = RGB(250, 252, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A" & noteRow & ":P" & noteRow).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 205, 230)
    End With
    
    ' ── 8. Navigation Button on SYNOPSIS ──
    Dim btn As Shape
    On Error Resume Next
    Set btn = ws.Shapes("btnSynopsisReturn")
    On Error GoTo ErrorHandler
    
    If btn Is Nothing Then
        Set btn = ws.Shapes.AddShape(5, ws.Range("N3").Left, ws.Range("N3").Top, _
                                    (ws.Range("P3").Left + ws.Range("P3").Width) - ws.Range("N3").Left, 22)
        With btn
            .Name = "btnSynopsisReturn"
            .Fill.ForeColor.RGB = RGB(32, 96, 33) ' UI_GREEN_DARK
            .Fill.Solid
            .Line.Visible = msoFalse
            With .TextFrame2
                .TextRange.Text = "← BACK TO QUOTE"
                .TextRange.Font.Size = 9
                .TextRange.Font.Bold = True
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .VerticalAnchor = msoAnchorMiddle
                .HorizontalAnchor = msoAnchorCenter
            End With
            .OnAction = "modUI.UI_ReturnToQuote"
        End With
    Else
        btn.Left = ws.Range("N3").Left
        btn.Top = ws.Range("N3").Top
        btn.Width = (ws.Range("P3").Left + ws.Range("P3").Width) - ws.Range("N3").Left
        btn.Height = 22
    End If
    
    ' ── 9. Page Setup for Clean Printing ──
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .PrintArea = "$A$1:$P$" & noteRow
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .CenterHorizontally = True
        .CenterVertically = False
    End With
    Application.PrintCommunication = True
    On Error GoTo ErrorHandler
    
    ' Ensure unlocked cells for user safety
    ws.Range("A1:P" & (noteRow + 5)).Locked = False
    
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

ErrorHandler:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Err.Raise Err.Number, "BuildSynopsisSheet", Err.Description
End Sub

Public Sub PopulateSynopsisReport()
    BuildSynopsisSheet
End Sub

Public Function GetSynopsisSetting(ByVal keyName As String, ByVal defaultValue As String) As String
    On Error Resume Next
    Dim wsSet As Worksheet
    Dim tblSet As ListObject
    Dim rObj As ListRow
    Set wsSet = ThisWorkbook.Worksheets("SETTINGS")
    If wsSet Is Nothing Then
        GetSynopsisSetting = defaultValue
        Exit Function
    End If
    Set tblSet = wsSet.ListObjects("tblSettings")
    If tblSet Is Nothing Then
        GetSynopsisSetting = defaultValue
        Exit Function
    End If
    For Each rObj In tblSet.ListRows
        If StrComp(Trim$(CStr(rObj.Range.Cells(1, tblSet.ListColumns("SettingKey").Index).Value)), keyName, vbTextCompare) = 0 Then
            Dim sVal As String
            sVal = Trim$(CStr(rObj.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value))
            If Len(sVal) > 0 Then
                GetSynopsisSetting = sVal
            Else
                GetSynopsisSetting = defaultValue
            End If
            Exit Function
        End If
    Next rObj
    GetSynopsisSetting = defaultValue
End Function

Public Sub SetSynopsisPresentationSettings(Optional ByVal inchMode As String = "", Optional ByVal pipeMode As String = "")
    On Error Resume Next
    Dim wsSet As Worksheet
    Dim tblSet As ListObject
    Dim rObj As ListRow
    Dim foundInch As Boolean
    Dim foundPipe As Boolean
    
    Set wsSet = ThisWorkbook.Worksheets("SETTINGS")
    If wsSet Is Nothing Then Exit Sub
    Set tblSet = wsSet.ListObjects("tblSettings")
    If tblSet Is Nothing Then Exit Sub
    
    If Len(inchMode) > 0 Then
        For Each rObj In tblSet.ListRows
            If StrComp(Trim$(CStr(rObj.Range.Cells(1, tblSet.ListColumns("SettingKey").Index).Value)), "SynopsisInchDisplay", vbTextCompare) = 0 Then
                rObj.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value = inchMode
                foundInch = True
                Exit For
            End If
        Next rObj
        If Not foundInch Then
            Set rObj = tblSet.ListRows.Add
            rObj.Range.Cells(1, tblSet.ListColumns("SettingKey").Index).Value = "SynopsisInchDisplay"
            rObj.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value = inchMode
            rObj.Range.Cells(1, tblSet.ListColumns("DataType").Index).Value = "Text"
            rObj.Range.Cells(1, tblSet.ListColumns("Description").Index).Value = "Synopsis Column C format (Nominal trade size vs Exact converted inches)"
            rObj.Range.Cells(1, tblSet.ListColumns("DefaultValue").Index).Value = "Nominal"
            rObj.Range.Cells(1, tblSet.ListColumns("Editable").Index).Value = "Yes"
            rObj.Range.Cells(1, tblSet.ListColumns("Category").Index).Value = "Presentation"
        End If
    End If
    
    If Len(pipeMode) > 0 Then
        For Each rObj In tblSet.ListRows
            If StrComp(Trim$(CStr(rObj.Range.Cells(1, tblSet.ListColumns("SettingKey").Index).Value)), "SynopsisPipeCountDisplay", vbTextCompare) = 0 Then
                rObj.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value = pipeMode
                foundPipe = True
                Exit For
            End If
        Next rObj
        If Not foundPipe Then
            Set rObj = tblSet.ListRows.Add
            rObj.Range.Cells(1, tblSet.ListColumns("SettingKey").Index).Value = "SynopsisPipeCountDisplay"
            rObj.Range.Cells(1, tblSet.ListColumns("SettingValue").Index).Value = pipeMode
            rObj.Range.Cells(1, tblSet.ListColumns("DataType").Index).Value = "Text"
            rObj.Range.Cells(1, tblSet.ListColumns("Description").Index).Value = "Synopsis Column I format (Dispatch whole joints vs Equivalent12M 2 decimals)"
            rObj.Range.Cells(1, tblSet.ListColumns("DefaultValue").Index).Value = "Dispatch"
            rObj.Range.Cells(1, tblSet.ListColumns("Editable").Index).Value = "Yes"
            rObj.Range.Cells(1, tblSet.ListColumns("Category").Index).Value = "Presentation"
        End If
    End If
    
    BuildSynopsisSheet
End Sub

Public Function RunSynopsisSelfTests() As String
    On Error GoTo TestFailure
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsQ As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    
    Set wb = ThisWorkbook
    Set wsQ = wb.Worksheets("QUOTE")
    Set wsLines = wb.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    ' 1. Rebuild Synopsis Sheet with default presentation settings
    SetSynopsisPresentationSettings "Nominal", "Dispatch"
    Set ws = wb.Worksheets("SYNOPSIS")
    
    If ws Is Nothing Then
        RunSynopsisSelfTests = "FAIL: SYNOPSIS worksheet does not exist"
        Exit Function
    End If
    
    ' If no quotation lines in active quote, populate golden quote for self-test
    If ws.Range("A12").MergeCells Then
        modSelfTest.PopulateGoldenReferenceQuote
        SetSynopsisPresentationSettings "Nominal", "Dispatch"
        Set ws = wb.Worksheets("SYNOPSIS")
    End If
    
    ' 2. Verify Title and Header
    If InStr(1, CStr(ws.Range("A1").Value), "SYNOPSIS", vbTextCompare) = 0 Then
        RunSynopsisSelfTests = "FAIL: Cell A1 is not SYNOPSIS title"
        Exit Function
    End If
    
    If ws.Range("A10").Value <> "Sr. No." Then
        RunSynopsisSelfTests = "FAIL: Column header A10 is not Sr. No."
        Exit Function
    End If
    
    If ws.Range("B10").Value <> "ID" Or ws.Range("B11").Value <> "MM" Then
        RunSynopsisSelfTests = "FAIL: Column header B10:B11 is not ID MM"
        Exit Function
    End If
    
    If ws.Range("D10").Value <> "OD " Or ws.Range("D11").Value <> "MM" Then
        RunSynopsisSelfTests = "FAIL: Column header D10:D11 is not OD MM"
        Exit Function
    End If
    
    ' 3. Verify Page Setup
    If ws.PageSetup.FitToPagesWide <> 1 Then
        RunSynopsisSelfTests = "FAIL: PageSetup FitToPagesWide is not 1"
        Exit Function
    End If
    
    ' 4. Verify Column M and P Precision Formulas when quotation lines exist
    If tblLines.ListRows.Count > 0 Then
        If InStr(1, ws.Cells(12, 13).Formula, "ROUNDUP", vbTextCompare) > 0 Or InStr(1, ws.Cells(12, 13).Formula, "L12/K12", vbTextCompare) = 0 Then
            RunSynopsisSelfTests = "FAIL: Column M formula is not unrounded =L12/K12"
            Exit Function
        End If
        If InStr(1, ws.Cells(12, 13).NumberFormat, "0.00", vbTextCompare) = 0 Then
            RunSynopsisSelfTests = "FAIL: Column M NumberFormat does not show 2 decimal places"
            Exit Function
        End If
        If InStr(1, ws.Cells(12, 16).Formula, "ROUNDUP", vbTextCompare) > 0 Or InStr(1, ws.Cells(12, 16).Formula, "12*J12/1000", vbTextCompare) = 0 Then
            RunSynopsisSelfTests = "FAIL: Column P formula is not unrounded =12*J12/1000"
            Exit Function
        End If
        If InStr(1, ws.Cells(12, 16).NumberFormat, "0.00", vbTextCompare) = 0 Then
            RunSynopsisSelfTests = "FAIL: Column P NumberFormat does not show 2 decimal places"
            Exit Function
        End If
        
        ' 5. Verify Column C (Default Nominal) and Column I (Default Dispatch)
        If InStr(1, ws.Cells(12, 3).Formula, "ROUND(", vbTextCompare) = 0 Then
            RunSynopsisSelfTests = "FAIL: Default Column C is not =ROUND(D12/25.4,0)"
            Exit Function
        End If
        If InStr(1, ws.Cells(12, 9).Formula, "CEILING(", vbTextCompare) = 0 Then
            RunSynopsisSelfTests = "FAIL: Default Column I is not =CEILING(H12/12,1)"
            Exit Function
        End If
        
        ' 6. Verify Presentation Setting Toggle (Exact and Equivalent12M)
        SetSynopsisPresentationSettings "Exact", "Equivalent12M"
        If InStr(1, ws.Cells(12, 3).Formula, "D12/25.4", vbTextCompare) = 0 Or InStr(1, ws.Cells(12, 3).Formula, "ROUND(", vbTextCompare) > 0 Then
            RunSynopsisSelfTests = "FAIL: Column C Exact setting failed"
            Exit Function
        End If
        If InStr(1, ws.Cells(12, 9).Formula, "H12/12", vbTextCompare) = 0 Or InStr(1, ws.Cells(12, 9).Formula, "CEILING(", vbTextCompare) > 0 Then
            RunSynopsisSelfTests = "FAIL: Column I Equivalent12M setting failed"
            Exit Function
        End If
        
        ' Restore default presentation settings
        SetSynopsisPresentationSettings "Nominal", "Dispatch"
    End If
    
    RunSynopsisSelfTests = "PASS"
    Exit Function

TestFailure:
    RunSynopsisSelfTests = "FAIL: " & Err.Description
End Function

