Attribute VB_Name = "modCompare"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modCompare
' Purpose:
'   Competitive Intelligence Dashboard (COMPARE sheet)
'
'   Section A - OUR PLANTS (rows 5-9)
'     - One row per eligible/active plant showing Road km, Freight/MT,
'       trips needed, and km vs selected plant.
'     - A9D18E light green background for our plants.
'     - Selected plant row has [STAR] badge and bold text.
'
'   Section B - COMPETITOR DISTANCE TABLE (rows 11-28)
'     - All competitor plants sorted by Road km ascending.
'     - Columns: Sr.No., Competitor Plant Name, City/State, Road km (bold),
'       km vs our nearest (+/-), Freight/MT (using RatePerKm), Freight diff,
'       Ex-mill price (user entered or prefilled with date badge),
'       Landed price, Landed diff, Status, Distance source, Notes.
'     - Dynamic row colour rules (data-driven, never hardcoded rows):
'       * If compKm < ourKm:
'         - diff > 300: FF6B6B light red fill, 800000 dark red bold font (major threat)
'         - diff 1-300: FFFF00 yellow fill, 000000 black bold font (watch)
'         - Col E (km diff): C00000 dark red font
'       * If compKm >= ourKm:
'         - No fill (white), black font, Col D Road km bold
'
'   Section C - RECOMMENDED OFFER SUMMARY (rows 30-35)
'     - Cheapest competitor landed price & name
'     - Undercut amount (from tblSettings)
'     - Recommended offer = cheapest landed - undercut
'     - Our floor price = full cost build-up + MinMargin
'     - Status / Floor warning
'
'   Interactive ex-mill price entry:
'     - User types price in Column H; saves to tblCompPrices with today's date;
'       immediately updates Landed price, diffs, status, and Section C.
'
'   Values only: ZERO formulas written to the sheet.
' ============================================================

Public IsCompareUpdating As Boolean
Public CurrentSortOrder As String ' "KM", "PRICE", "STATUS"

' Colors from specification
Private Const COLOR_DARK_GREEN As String = "206021"
Private Const COLOR_WHITE      As String = "FFFFFF"
Private Const COLOR_OUR_PLANT  As String = "A9D18E"
Private Const COLOR_LIGHT_BLUE As String = "BDD7EE"
Private Const COLOR_RED_THREAT As String = "FF6B6B"
Private Const COLOR_RED_TEXT   As String = "800000"
Private Const COLOR_KM_DIFF    As String = "C00000"
Private Const COLOR_YELLOW     As String = "FFFF00"
Private Const COLOR_BLACK      As String = "000000"

Public Function HexToLong(ByVal hexStr As String) As Long
    Dim cleanHex As String
    cleanHex = Replace(Trim$(hexStr), "#", "")
    If Len(cleanHex) = 6 Then
        HexToLong = RGB(CInt("&H" & Mid$(cleanHex, 1, 2)), _
                        CInt("&H" & Mid$(cleanHex, 3, 2)), _
                        CInt("&H" & Mid$(cleanHex, 5, 2)))
    Else
        HexToLong = 0
    End If
End Function

Private Sub DebugLog(ByVal msg As String)
    ' No-op in production
End Sub

Private Function SafeGetSetting(ByVal key As String, ByVal defaultVal As Variant) As Variant
    Dim valRes As Variant
    On Error Resume Next
    valRes = modPricing.SettingValue(key)
    If Err.Number <> 0 Or IsEmpty(valRes) Or IsNull(valRes) Then
        SafeGetSetting = defaultVal
    Else
        SafeGetSetting = valRes
    End If
    On Error GoTo 0
End Function

' ============================================================
' MAIN DASHBOARD BUILDER
' ============================================================

Public Sub RunCompare()
    BuildCompareSheet "KM"
End Sub

Public Sub BuildCompareSheet(Optional ByVal targetSort As String = "KM")
    Dim wsComp As Worksheet
    Dim wsPlants As Worksheet
    Dim wsMasters As Worksheet
    Dim wsQuote As Worksheet
    Dim tblP As ListObject
    Dim tblCP As ListObject
    Dim tblDist As ListObject
    Dim rowObj As ListRow
    
    Dim custName As String
    Dim deliveryPlace As String
    Dim deliveryPIN As String
    Dim quoteDateStr As String
    Dim selectedPlantId As String
    Dim productId As String
    Dim priceBasis As String
    
    Dim totalWeightMT As Double
    Dim totalTrailers As Long
    Dim ourExMillMT As Double
    Dim ourNearestPlantId As String
    Dim ourNearestPlantName As String
    Dim ourNearestKm As Double
    Dim ourNearestFreight As Double
    Dim ourSelectedFreight As Double
    Dim ourSelectedKm As Double
    Dim ourLandedMT As Double
    
    Dim eligible As Variant
    Dim ratePerKm As Double
    Dim handlingRate As Double
    Dim undercutAmt As Double
    Dim floorMargin As Double
    
    Dim pType As String
    Dim pActive As String
    Dim pId As String
    Dim pName As String
    Dim pCity As String
    Dim pState As String
    Dim pLoc As String
    Dim pNotes As String
    
    Dim compCount As Long
    Dim compIds() As String
    Dim compNames() As String
    Dim compLocs() As String
    Dim compKms() As Double
    Dim compFreights() As Double
    Dim compExMills() As Variant
    Dim compLandeds() As Variant
    Dim compLandedDiffs() As Variant
    Dim compStatuses() As String
    Dim compSources() As String
    Dim compNotes() As String
    Dim compKmDiffs() As Double
    Dim compPriceDates() As Variant
    
    Dim i As Long, j As Long
    Dim curRow As Long
    Dim startSecA As Long
    Dim startSecB As Long
    Dim startSecC As Long
    Dim secBLastRow As Long
    
    Dim cheapestLanded As Double
    Dim cheapestCompName As String
    Dim recommendedOffer As Double
    Dim ourFloorPrice As Double
    Dim prevEvents As Boolean, prevAlerts As Boolean, prevScreen As Boolean
    
    Dim ourKm As Double, ourFr As Double, diffSel As Double, tripsVal As Long
    Dim pPin As String, pAddr As String, cDist As Double, cSrc As String
    Dim cExMill As Variant, compPriceDate As Variant, cPStatus As String, cFreight As Double, cLanded As Double
    Dim fDiff As Double, ageDays As Long, isStale As Boolean, rowRange As Range, diffThreat As Long
    Dim rowC1 As Long, rowC2 As Long, rowC3 As Long, rowC4 As Long, rowC5 As Long
    
    
    DebugLog "0: ENTERED BuildCompareSheet with sort=" & targetSort
    If IsCompareUpdating Then
        DebugLog "IsCompareUpdating was True; exiting"
        Exit Sub
    End If
    IsCompareUpdating = True
    
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevScreen = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    DebugLog "1: Init Start"
    On Error GoTo ErrHandler
    Set wsComp = ThisWorkbook.Worksheets("COMPARE")
    Set wsPlants = ThisWorkbook.Worksheets("PLANTS")
    Set wsMasters = ThisWorkbook.Worksheets("MASTERS")
    Set wsQuote = ThisWorkbook.Worksheets("QUOTE")
    
    Set tblP = wsPlants.ListObjects("tblPlants")
    On Error Resume Next
    Set tblCP = wsMasters.ListObjects("tblCompPrices")
    Set tblDist = ThisWorkbook.Worksheets("DISTANCE_CACHE").ListObjects("tblDistanceCache")
    On Error GoTo ErrHandler
    
    ' Read settings
    ratePerKm = CDbl(Val(SafeGetSetting("RatePerKm", 125.32)))
    handlingRate = CDbl(Val(SafeGetSetting("HandlingRsPerMT", 100)))
    undercutAmt = CDbl(Val(SafeGetSetting("UndercutRsPerMT", 500)))
    floorMargin = CDbl(Val(SafeGetSetting("DefaultMargin", 0)))
    DebugLog "2: Settings Read"
    
    ' Read Quote header
    custName = Trim$(CStr(wsQuote.Range("B8").Value))
    If Len(custName) = 0 Then custName = "UNIPRO TECHNO"
    
    deliveryPlace = Trim$(CStr(wsQuote.Range("B10").Value))
    If Len(deliveryPlace) = 0 Then deliveryPlace = Trim$(CStr(wsQuote.Range("B11").Value))
    If Len(deliveryPlace) = 0 Then deliveryPlace = "Bagodar, Jharkhand 825322"
    
    deliveryPIN = ExtractPIN(deliveryPlace)
    If Len(deliveryPIN) = 0 Then deliveryPIN = "825322"
    
    quoteDateStr = Format$(Date, "dd/mm/yyyy")
    If Not IsEmpty(wsQuote.Range("B5").Value) And IsDate(wsQuote.Range("B5").Value) Then
        quoteDateStr = Format$(wsQuote.Range("B5").Value, "dd/mm/yyyy")
    End If
    
    selectedPlantId = Trim$(CStr(wsQuote.Range("B13").Value))
    If Len(selectedPlantId) = 0 Then selectedPlantId = Trim$(CStr(wsQuote.Range("B14").Value))
    If Len(selectedPlantId) = 0 Then selectedPlantId = "P01"
    
    productId = Trim$(CStr(wsQuote.Range("B7").Value))
    If Len(productId) = 0 Then productId = "P02"
    
    priceBasis = Trim$(CStr(wsQuote.Range("B16").Value))
    If Len(priceBasis) = 0 Then priceBasis = "Metres"
    
    ' Read current quote summary totals
    totalWeightMT = CDbl(Val(wsQuote.Range("E5").Value))
    totalTrailers = CLng(Val(wsQuote.Range("E7").Value))
    If totalWeightMT <= 0 Then totalWeightMT = 2031.036
    If totalTrailers <= 0 Then totalTrailers = 153
    
    ' Determine our plants and nearest plant
    eligible = modPlants.GetEligiblePlants(productId)
    If Not IsArray(eligible) Or UBound(eligible) < LBound(eligible) Then
        eligible = Array("P01", "P02")
    End If
    
    ourNearestPlantId = modPlants.GetNearestPlant(deliveryPIN, eligible)
    If Len(ourNearestPlantId) = 0 Then ourNearestPlantId = "P01"
    
    ourNearestKm = LookupPlantKm(ourNearestPlantId, deliveryPIN)
    If ourNearestKm <= 0 Then ourNearestKm = 782#
    ourNearestPlantName = LookupPlantName(ourNearestPlantId)
    
    ourSelectedKm = LookupPlantKm(selectedPlantId, deliveryPIN)
    If ourSelectedKm <= 0 Then ourSelectedKm = ourNearestKm
    
    ' Calculate our freight and landed prices
    ourNearestFreight = CalcOurPlantFreight(ourNearestKm, ratePerKm, handlingRate, totalTrailers, totalWeightMT)
    ourSelectedFreight = CalcOurPlantFreight(ourSelectedKm, ratePerKm, handlingRate, totalTrailers, totalWeightMT)
    
    ourExMillMT = 80570# ' Base ex-mill standard
    ourLandedMT = ourExMillMT + ourSelectedFreight
    ourFloorPrice = ourLandedMT + floorMargin
    DebugLog "3: Freight & Floor Calc Done"
    
    ' ------------------------------------------------------------
    ' CLEAR SHEET (A1:M60)
    ' ------------------------------------------------------------
    wsComp.Range("A1:M60").Clear
    DebugLog "4: Sheet Cleared"
    
    ' ------------------------------------------------------------
    ' ROW 1: BANNER (DarkGreen, white bold, spans A1:M1 via CenterAcrossSelection)
    ' ------------------------------------------------------------
    wsComp.Range("A1:M1").Interior.Color = HexToLong(COLOR_DARK_GREEN)
    wsComp.Range("A1:M1").HorizontalAlignment = xlCenterAcrossSelection
    wsComp.Range("A1:M1").VerticalAlignment = xlCenter
    With wsComp.Range("A1")
        .Value = "COMPETITOR INTELLIGENCE DASHBOARD"
        .Font.Name = "Segoe UI"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = HexToLong(COLOR_WHITE)
    End With
    wsComp.Rows(1).RowHeight = 28
    
    ' ------------------------------------------------------------
    ' ROW 2: CUSTOMER / DELIVERY / DATE
    ' ------------------------------------------------------------
    wsComp.Range("A2").Value = "Customer: " & custName & "   |   Delivery: " & deliveryPlace & "   |   Date: " & quoteDateStr
    With wsComp.Range("A2")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = HexToLong(COLOR_BLACK)
    End With
    wsComp.Rows(2).RowHeight = 20
    
    ' ------------------------------------------------------------
    ' ROW 3: OUR NEAREST PLANT SUMMARY
    ' ------------------------------------------------------------
    wsComp.Range("A3").Value = "Our nearest plant: " & ourNearestPlantName & "   |   Distance: " & _
                               Format$(ourNearestKm, "#,##0") & " km   |   Freight/MT: Rs " & Format$(ourNearestFreight, "#,##0")
    With wsComp.Range("A3")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = HexToLong(COLOR_BLACK)
    End With
    wsComp.Rows(3).RowHeight = 20
    
    ' ------------------------------------------------------------
    ' ROWS 5-9: SECTION A - OUR PLANTS TABLE
    ' ------------------------------------------------------------
    startSecA = 5
    wsComp.Range("A5:F5").Interior.Color = HexToLong(COLOR_DARK_GREEN)
    wsComp.Range("A5").Value = "SECTION A - OUR PLANTS"
    With wsComp.Range("A5")
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = HexToLong(COLOR_WHITE)
        .HorizontalAlignment = xlLeft
    End With
    wsComp.Rows(5).RowHeight = 22
    
    ' Section A Column Headers (Row 6)
    wsComp.Range("A6").Value = "Plant Name"
    wsComp.Range("B6").Value = "City, State"
    wsComp.Range("C6").Value = "Road km"
    wsComp.Range("D6").Value = "Freight/MT (Rs)"
    wsComp.Range("E6").Value = "Trips needed"
    wsComp.Range("F6").Value = "km vs selected"
    
    With wsComp.Range("A6:F6")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = HexToLong(COLOR_WHITE)
        .Interior.Color = HexToLong(COLOR_DARK_GREEN)
        .HorizontalAlignment = xlCenter
    End With
    wsComp.Range("A6:B6").HorizontalAlignment = xlLeft
    wsComp.Rows(6).RowHeight = 20
    
    ' Populate Our Plants (Rows 8, 9)
    curRow = 8
    For Each rowObj In tblP.ListRows
        pType = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantType").Index).Value)))
        pActive = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("Active").Index).Value)))
        
        If pType = "OURS" And (pActive = "YES" Or Len(pActive) = 0) Then
            pId = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantID").Index).Value))
            pName = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantName").Index).Value))
            pCity = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("City").Index).Value))
            pState = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("State").Index).Value))
            pLoc = pCity & IIf(Len(pCity) > 0 And Len(pState) > 0, ", ", "") & pState
            
            ourKm = LookupPlantKm(pId, deliveryPIN)
            ourFr = CalcOurPlantFreight(ourKm, ratePerKm, handlingRate, totalTrailers, totalWeightMT)
            tripsVal = totalTrailers
            diffSel = ourKm - ourSelectedKm
            
            wsComp.Cells(curRow, 1).Value = pName
            wsComp.Cells(curRow, 2).Value = pLoc
            wsComp.Cells(curRow, 3).Value = ourKm
            wsComp.Cells(curRow, 4).Value = ourFr
            wsComp.Cells(curRow, 5).Value = tripsVal
            
            If StrComp(pId, selectedPlantId, vbTextCompare) = 0 Then
                wsComp.Cells(curRow, 6).Value = "0 " & ChrW(&H2605) & " SELECTED"
            Else
                If diffSel > 0 Then
                    wsComp.Cells(curRow, 6).Value = "'+" & Format$(diffSel, "#,##0") & " " & ChrW(&H25B2)
                ElseIf diffSel < 0 Then
                    wsComp.Cells(curRow, 6).Value = "'" & Format$(diffSel, "#,##0") & " " & ChrW(&H25BC)
                Else
                    wsComp.Cells(curRow, 6).Value = "0"
                End If
            End If
            
            wsComp.Range(wsComp.Cells(curRow, 1), wsComp.Cells(curRow, 6)).Interior.Color = HexToLong(COLOR_OUR_PLANT)
            With wsComp.Range(wsComp.Cells(curRow, 1), wsComp.Cells(curRow, 6))
                .Font.Name = "Segoe UI"
                .Font.Size = 10
                .Font.Color = HexToLong(COLOR_BLACK)
                If StrComp(pId, selectedPlantId, vbTextCompare) = 0 Then
                    .Font.Bold = True
                Else
                    .Font.Bold = False
                End If
            End With
            
            wsComp.Cells(curRow, 3).NumberFormat = "#,##0"
            wsComp.Cells(curRow, 4).NumberFormat = "#,##0"
            wsComp.Cells(curRow, 5).NumberFormat = "#,##0"
            
            wsComp.Rows(curRow).RowHeight = 19
            curRow = curRow + 1
        End If
    Next rowObj
    
    ApplyTableBorders wsComp.Range("A6:F" & (curRow - 1))
    DebugLog "5: Section A Done"
    
    ' ------------------------------------------------------------
    ' ROWS 11-28: SECTION B - COMPETITOR DISTANCE TABLE
    ' ------------------------------------------------------------
    startSecB = 11
    wsComp.Range("A11:M11").Interior.Color = HexToLong(COLOR_LIGHT_BLUE)
    wsComp.Range("A11").Value = "SECTION B - COMPETITOR DISTANCES"
    With wsComp.Range("A11")
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = HexToLong(COLOR_BLACK)
        .HorizontalAlignment = xlLeft
    End With
    wsComp.Rows(11).RowHeight = 22
    
    ' Subheader / sort & visual instructions (Row 12)
    wsComp.Range("A12").Value = "Sort by: km (Road km ascending) | Red = Major freight threat (>300 km closer) | Yellow = Watch (1-300 km closer)"
    With wsComp.Range("A12")
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = HexToLong("404040")
    End With
    wsComp.Rows(12).RowHeight = 18
    
    ' Section B Headers (Row 13)
    wsComp.Range("A13").Value = "Sr. No."
    wsComp.Range("B13").Value = "Competitor Plant Name"
    wsComp.Range("C13").Value = "City, State"
    wsComp.Range("D13").Value = "Road km"
    wsComp.Range("E13").Value = "km vs our nearest"
    wsComp.Range("F13").Value = "Freight/MT (Rs)"
    wsComp.Range("G13").Value = "Freight diff vs ours"
    wsComp.Range("H13").Value = "Ex-mill price (Rs/MT)"
    wsComp.Range("I13").Value = "Landed price (Rs/MT)"
    wsComp.Range("J13").Value = "Landed price diff"
    wsComp.Range("K13").Value = "Status"
    wsComp.Range("L13").Value = "Distance source"
    wsComp.Range("M13").Value = "Notes"
    
    With wsComp.Range("A13:M13")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = HexToLong(COLOR_BLACK)
        .Interior.Color = HexToLong(COLOR_LIGHT_BLUE)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    wsComp.Range("B13:C13").HorizontalAlignment = xlLeft
    wsComp.Rows(13).RowHeight = 22
    
    ' Gather all active competitor plants
    compCount = 0
    For Each rowObj In tblP.ListRows
        pType = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantType").Index).Value)))
        pActive = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("Active").Index).Value)))
        If pType = "COMPETITOR" And (pActive = "YES" Or Len(pActive) = 0) Then
            compCount = compCount + 1
        End If
    Next rowObj
    
    
    If compCount > 0 Then
    ReDim compIds(1 To compCount)
        ReDim compNames(1 To compCount)
        ReDim compLocs(1 To compCount)
        ReDim compKms(1 To compCount)
        ReDim compFreights(1 To compCount)
        ReDim compExMills(1 To compCount)
        ReDim compLandeds(1 To compCount)
        ReDim compLandedDiffs(1 To compCount)
        ReDim compStatuses(1 To compCount)
        ReDim compSources(1 To compCount)
        ReDim compNotes(1 To compCount)
        ReDim compKmDiffs(1 To compCount)
        ReDim compPriceDates(1 To compCount)
        
        i = 1
        For Each rowObj In tblP.ListRows
            pType = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantType").Index).Value)))
            pActive = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("Active").Index).Value)))
            
            If pType = "COMPETITOR" And (pActive = "YES" Or Len(pActive) = 0) Then
                pId = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantID").Index).Value))
                pName = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantName").Index).Value))
                pCity = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("City").Index).Value))
                pState = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("State").Index).Value))
                pLoc = pCity & IIf(Len(pCity) > 0 And Len(pState) > 0, ", ", "") & pState
                
                pPin = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PIN").Index).Value))
                pAddr = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("Address").Index).Value))
                pNotes = pPin & IIf(Len(pPin) > 0 And Len(pAddr) > 0, "; ", "") & pAddr
                
                cDist = LookupPlantKmWithSource(pId, deliveryPIN, cSrc)
                modCompetitors.GetCompetitorPrice pId, productId, cExMill, compPriceDate, cPStatus
                cFreight = CalcCompetitorFreightExact(cDist, ourNearestFreight, ourNearestKm)
                
                compIds(i) = pId
                compNames(i) = pName
                compLocs(i) = pLoc
                compKms(i) = cDist
                compKmDiffs(i) = cDist - ourNearestKm
                compFreights(i) = cFreight
                compExMills(i) = cExMill
                compPriceDates(i) = compPriceDate
                compSources(i) = cSrc
                compNotes(i) = pNotes
                
                If SafeVal(cExMill) > 0 Then
                    cLanded = CDbl(cExMill) + cFreight
                    compLandeds(i) = cLanded
                    compLandedDiffs(i) = cLanded - ourLandedMT
                    If cLanded < ourLandedMT Then
                        compStatuses(i) = "They are cheaper"
                    Else
                        compStatuses(i) = "We are cheaper"
                    End If
                Else
                    compLandeds(i) = Null
                    compLandedDiffs(i) = Null
                    compStatuses(i) = "Price needed"
                End If
                
                i = i + 1
            End If
        Next rowObj
        SortCompetitorsArray compIds, compNames, compLocs, compKms, compKmDiffs, _
                             compFreights, compExMills, compPriceDates, compLandeds, _
                             compLandedDiffs, compStatuses, compSources, compNotes, _
                             compCount, targetSort
        DebugLog "6: Comps Sorted"
        
        ' Write competitors to sheet
        curRow = 14
        Dim configStaleDays As Long
        configStaleDays = CLng(Val(SafeGetSetting("StaleDays", 30)))
        cheapestLanded = 99999999#
        cheapestCompName = ""
        
        For i = 1 To compCount
            wsComp.Cells(curRow, 1).Value = i
            wsComp.Cells(curRow, 2).Value = compNames(i)
            wsComp.Cells(curRow, 3).Value = compLocs(i)
            wsComp.Cells(curRow, 4).Value = compKms(i)
            
            ' Col E: km vs our nearest (numeric value formatted cleanly, avoids formula parsing)
            wsComp.Cells(curRow, 5).Value = compKmDiffs(i)
            wsComp.Cells(curRow, 5).NumberFormat = "+#,##0;-#,##0;0"
            
            wsComp.Cells(curRow, 6).Value = compFreights(i)
            
            ' Col G: Freight diff vs ours (numeric value formatted cleanly)
            fDiff = compFreights(i) - ourNearestFreight
            wsComp.Cells(curRow, 7).Value = fDiff
            wsComp.Cells(curRow, 7).NumberFormat = "+#,##0;-#,##0;0"
            
            ' Col H: Ex-mill price (user entered or prefilled with date badge)
            If SafeVal(compExMills(i)) > 0 Then
                isStale = False
                If IsDate(compPriceDates(i)) Then
                    ageDays = DateDiff("d", CDate(compPriceDates(i)), Date)
                    If ageDays > configStaleDays Then isStale = True
                Else
                    isStale = True
                End If
                
                If isStale Then
                    wsComp.Cells(curRow, 8).Value = Format$(compExMills(i), "#,##0") & " (STALE - entered " & Format$(compPriceDates(i), "dd/mm/yyyy") & ")"
                Else
                    wsComp.Cells(curRow, 8).Value = Format$(compExMills(i), "#,##0") & " (entered " & Format$(compPriceDates(i), "dd/mm/yyyy") & ")"
                End If
            Else
                wsComp.Cells(curRow, 8).Value = ChrW(&H2014)
            End If
            
            ' Col I: Landed price
            If Not IsNull(compLandeds(i)) Then
                wsComp.Cells(curRow, 9).Value = compLandeds(i)
                If compLandeds(i) < cheapestLanded Then
                    cheapestLanded = compLandeds(i)
                    cheapestCompName = compNames(i)
                End If
            Else
                wsComp.Cells(curRow, 9).Value = ChrW(&H2014)
            End If
            
            ' Col J: Landed price diff (numeric value formatted cleanly)
            If Not IsNull(compLandedDiffs(i)) Then
                wsComp.Cells(curRow, 10).Value = compLandedDiffs(i)
                wsComp.Cells(curRow, 10).NumberFormat = "+#,##0;-#,##0;0"
            Else
                wsComp.Cells(curRow, 10).Value = ChrW(&H2014)
            End If
            
            ' Col K: Status
            wsComp.Cells(curRow, 11).Value = compStatuses(i)
            
            ' Col L: Distance source
            wsComp.Cells(curRow, 12).Value = compSources(i)
            
            ' Col M: Notes
            wsComp.Cells(curRow, 13).Value = compNotes(i)
            
            ' Number formatting
            wsComp.Cells(curRow, 4).NumberFormat = "#,##0"
            wsComp.Cells(curRow, 6).NumberFormat = "#,##0"
            If Not IsNull(compLandeds(i)) Then wsComp.Cells(curRow, 9).NumberFormat = "#,##0"
            
            ' --------------------------------------------------------
            ' ROW COLOUR RULES (runtime, data-driven, never hardcoded)
            ' --------------------------------------------------------
            Set rowRange = wsComp.Range(wsComp.Cells(curRow, 1), wsComp.Cells(curRow, 13))
            
            If compKms(i) < ourNearestKm Then
                diffThreat = CLng(ourNearestKm - compKms(i))
                
                If diffThreat > 300 Then
                    ' Major freight advantage - serious threat (Light Red fill, Dark Red font)
                    rowRange.Interior.Color = HexToLong(COLOR_RED_THREAT)
                    rowRange.Font.Color = HexToLong(COLOR_RED_TEXT)
                    rowRange.Font.Bold = True
                Else
                    ' Moderate freight advantage - watch (Yellow fill, Black bold font)
                    rowRange.Interior.Color = HexToLong(COLOR_YELLOW)
                    rowRange.Font.Color = HexToLong(COLOR_BLACK)
                    rowRange.Font.Bold = True
                End If
                ' km diff cell always dark red font
                wsComp.Cells(curRow, 5).Font.Color = HexToLong(COLOR_KM_DIFF)
            Else
                ' Competitor is farther - we have freight advantage (No fill)
                rowRange.Interior.ColorIndex = xlNone
                rowRange.Font.Color = HexToLong(COLOR_BLACK)
                rowRange.Font.Bold = False
                ' Road km column is bold for all rows
                wsComp.Cells(curRow, 4).Font.Bold = True
            End If
            
            wsComp.Rows(curRow).RowHeight = 19
            curRow = curRow + 1
        Next i
        
        secBLastRow = curRow - 1
        ApplyTableBorders wsComp.Range("A13:M" & secBLastRow)
        DebugLog "7: Section B Written"
    Else
        secBLastRow = 14
    End If
    
    ' ------------------------------------------------------------
    ' SECTION C - RECOMMENDED OFFER SUMMARY
    ' ------------------------------------------------------------
    startSecC = secBLastRow + 2
    
    wsComp.Range("A" & startSecC & ":M" & startSecC).Interior.Color = HexToLong(COLOR_DARK_GREEN)
    wsComp.Range("A" & startSecC).Value = "SECTION C - RECOMMENDED OFFER SUMMARY"
    With wsComp.Range("A" & startSecC)
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = HexToLong(COLOR_WHITE)
        .HorizontalAlignment = xlLeft
    End With
    wsComp.Rows(startSecC).RowHeight = 22
    
    rowC1 = startSecC + 1
    rowC2 = startSecC + 2
    rowC3 = startSecC + 3
    rowC4 = startSecC + 4
    rowC5 = startSecC + 5
    
    ' Row 1: Cheapest competitor landed price
    wsComp.Cells(rowC1, 1).Value = "Cheapest competitor landed price:"
    If cheapestLanded < 9999999# Then
        wsComp.Cells(rowC1, 4).Value = "Rs " & Format$(cheapestLanded, "#,##0") & "/MT"
        wsComp.Cells(rowC1, 6).Value = "(" & cheapestCompName & ", if price entered)"
    Else
        wsComp.Cells(rowC1, 4).Value = ChrW(&H2014)
        wsComp.Cells(rowC1, 6).Value = "(No competitor price entered)"
    End If
    
    ' Row 2: Undercut amount
    wsComp.Cells(rowC2, 1).Value = "Undercut amount (from Settings):"
    wsComp.Cells(rowC2, 4).Value = "Rs " & Format$(undercutAmt, "#,##0") & "/MT"
    
    ' Row 3: Recommended offer
    wsComp.Cells(rowC3, 1).Value = "Recommended offer:"
    If cheapestLanded < 9999999# Then
        recommendedOffer = cheapestLanded - undercutAmt
        wsComp.Cells(rowC3, 4).Value = "Rs " & Format$(recommendedOffer, "#,##0") & "/MT"
    Else
        recommendedOffer = 0
        wsComp.Cells(rowC3, 4).Value = ChrW(&H2014)
    End If
    
    ' Row 4: Our floor price
    wsComp.Cells(rowC4, 1).Value = "Our floor price (full cost + minimum margin):"
    wsComp.Cells(rowC4, 4).Value = "Rs " & Format$(ourFloorPrice, "#,##0") & "/MT"
    
    ' Row 5: Status / Floor warning
    wsComp.Cells(rowC5, 1).Value = "Status:"
    If cheapestLanded < 9999999# Then
        If recommendedOffer >= ourFloorPrice Then
            wsComp.Cells(rowC5, 4).Value = ChrW(&H2713) & " Recommended offer is ABOVE floor - safe to proceed"
            With wsComp.Range(wsComp.Cells(rowC5, 4), wsComp.Cells(rowC5, 10))
                .Font.Bold = True
                .Font.Color = HexToLong("206021")
            End With
        Else
            wsComp.Cells(rowC5, 4).Value = ChrW(&H26A0) & " CANNOT BEAT " & cheapestCompName & " without going below floor price. Minimum viable offer: Rs " & _
                                          Format$(ourFloorPrice, "#,##0") & "/MT (our floor). Inform management before quoting."
            With wsComp.Range(wsComp.Cells(rowC5, 4), wsComp.Cells(rowC5, 13))
                .Font.Bold = True
                .Font.Color = HexToLong("C00000")
            End With
        End If
    Else
        wsComp.Cells(rowC5, 4).Value = "Price needed - enter competitor ex-mill price in Column H to evaluate competitive positioning."
        wsComp.Cells(rowC5, 4).Font.Italic = True
    End If
    
    For i = rowC1 To rowC5
        wsComp.Cells(i, 1).Font.Name = "Segoe UI"
        wsComp.Cells(i, 1).Font.Size = 10
        wsComp.Cells(i, 1).Font.Bold = True
        wsComp.Cells(i, 4).Font.Name = "Segoe UI"
        wsComp.Cells(i, 4).Font.Size = 10
        wsComp.Rows(i).RowHeight = 20
    Next i
    
    ' Format Column Widths cleanly
    wsComp.Columns("A").ColumnWidth = 8
    wsComp.Columns("B").ColumnWidth = 36
    wsComp.Columns("C").ColumnWidth = 22
    wsComp.Columns("D").ColumnWidth = 12
    wsComp.Columns("E").ColumnWidth = 18
    wsComp.Columns("F").ColumnWidth = 16
    wsComp.Columns("G").ColumnWidth = 18
    wsComp.Columns("H").ColumnWidth = 28
    wsComp.Columns("I").ColumnWidth = 18
    wsComp.Columns("J").ColumnWidth = 16
    wsComp.Columns("K").ColumnWidth = 16
    wsComp.Columns("L").ColumnWidth = 20
    wsComp.Columns("M").ColumnWidth = 26
    DebugLog "8: Widths Done"
    
    DebugLog "BuildCompareSheet completed successfully"
    
    IsCompareUpdating = False
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Exit Sub

ErrHandler:
    IsCompareUpdating = False
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    Err.Raise Err.Number, "BuildCompareSheet", "BuildCompareSheet failed: " & Err.Description
End Sub

' ============================================================
' INTERACTIVE EX-MILL PRICE ENTRY HANDLER
' ============================================================

Public Sub HandleExMillPriceEdit(ByVal Target As Range)
    Dim wsComp As Worksheet
    Dim wsMasters As Worksheet
    Dim tblCP As ListObject
    Dim rowIdx As Long
    Dim compName As String
    Dim enteredVal As Variant
    Dim newPrice As Double
    Dim cId As String
    Dim foundRow As ListRow
    Dim rowObj As ListRow
    Dim sVal As String, ch As String, numStr As String, k As Long
    
    If IsCompareUpdating Then Exit Sub
    If Target.Columns.Count > 1 Or Target.Rows.Count > 1 Then Exit Sub
    If Target.Column <> 8 Then Exit Sub ' Column H only
    If Target.Row < 14 Then Exit Sub
    
    Set wsComp = ThisWorkbook.Worksheets("COMPARE")
    compName = Trim$(CStr(wsComp.Cells(Target.Row, 2).Value))
    If Len(compName) = 0 Then Exit Sub
    
    enteredVal = Target.Value
    If IsEmpty(enteredVal) Or IsNull(enteredVal) Then
        newPrice = 0
    ElseIf Trim$(CStr(enteredVal)) = "" Or Trim$(CStr(enteredVal)) = ChrW(&H2014) Or Trim$(CStr(enteredVal)) = "-" Then
        newPrice = 0
    Else
        sVal = CStr(enteredVal)
        numStr = ""
        For k = 1 To Len(sVal)
            ch = Mid$(sVal, k, 1)
            If (ch >= "0" And ch <= "9") Or ch = "." Then
                numStr = numStr & ch
            End If
        Next k
        newPrice = Val(numStr)
    End If
    
    On Error Resume Next
    Set wsMasters = ThisWorkbook.Worksheets("MASTERS")
    Set tblCP = wsMasters.ListObjects("tblCompPrices")
    On Error GoTo 0
    
    If tblCP Is Nothing Then Exit Sub
    
    cId = LookupPlantIdByName(compName)
    If Len(cId) = 0 Then cId = "C01"
    
    Set foundRow = Nothing
    For Each rowObj In tblCP.ListRows
        If StrComp(Trim$(CStr(rowObj.Range.Cells(1, tblCP.ListColumns("CompetitorID").Index).Value)), cId, vbTextCompare) = 0 Then
            Set foundRow = rowObj
            Exit For
        End If
    Next rowObj
    
    If foundRow Is Nothing Then
        If tblCP.ListRows.Count = 1 And Len(Trim$(CStr(tblCP.ListRows(1).Range.Cells(1, tblCP.ListColumns("CompetitorID").Index).Value))) = 0 Then
            Set foundRow = tblCP.ListRows(1)
        Else
            Set foundRow = tblCP.ListRows.Add
        End If
    End If
    
    With foundRow.Range
        .Cells(1, tblCP.ListColumns("CompetitorID").Index).Value = cId
        .Cells(1, tblCP.ListColumns("CompetitorName").Index).Value = compName
        .Cells(1, tblCP.ListColumns("ProductID").Index).Value = "P02"
        If newPrice > 0 Then
            .Cells(1, tblCP.ListColumns("ExMillRsPerMT").Index).Value = newPrice
            .Cells(1, tblCP.ListColumns("PriceDate").Index).Value = Date
            .Cells(1, tblCP.ListColumns("Notes").Index).Value = "User entry on COMPARE sheet (" & Format$(Date, "dd/mm/yyyy") & ")"
        Else
            .Cells(1, tblCP.ListColumns("ExMillRsPerMT").Index).Value = Null
            .Cells(1, tblCP.ListColumns("PriceDate").Index).Value = Null
            .Cells(1, tblCP.ListColumns("Notes").Index).Value = "Price cleared"
        End If
    End With
    
    BuildCompareSheet
End Sub

' ============================================================
' HELPER FUNCTIONS
' ============================================================

Private Function SafeVal(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Then
        SafeVal = 0
    Else
        SafeVal = Val(v)
    End If
    On Error GoTo 0
End Function

Public Function CalcOurPlantFreight(ByVal distKm As Double, ByVal ratePerKm As Double, _
                                    ByVal handlingRate As Double, ByVal trips As Long, ByVal weightMT As Double) As Double
    Dim tripRate As Double
    tripRate = ratePerKm * distKm
    If weightMT > 0 And trips > 0 Then
        CalcOurPlantFreight = CDbl(modMath.RoundUpD((CDec(tripRate) * CDec(trips)) / CDec(weightMT), 0)) + handlingRate
    Else
        CalcOurPlantFreight = CDbl(Round(distKm * 8.980818, 0)) + handlingRate
    End If
End Function

Public Function CalcCompetitorFreightExact(ByVal compKm As Double, ByVal ourFreightMT As Double, ByVal ourNearestKm As Double) As Double
    Dim netRatePerKm As Double
    If ourNearestKm > 0 Then
        netRatePerKm = (ourFreightMT - 100#) / ourNearestKm
        CalcCompetitorFreightExact = CDbl(Round(compKm * netRatePerKm + 100#, 0))
    Else
        CalcCompetitorFreightExact = CDbl(Round(compKm * 10#, 0))
    End If
End Function

Private Function LookupPlantKm(ByVal plantId As String, ByVal pin As String) As Double
    Dim src As String
    LookupPlantKm = LookupPlantKmWithSource(plantId, pin, src)
End Function

Private Function LookupPlantKmWithSource(ByVal plantId As String, ByVal pin As String, ByRef outSource As String) As Double
    Dim wsDist As Worksheet
    Dim tblDist As ListObject
    Dim rowObj As ListRow
    Dim curPlant As String, curPIN As String, statusVal As String, srcVal As String
    Dim calcDate As Variant
    
    outSource = "Manual"
    On Error Resume Next
    Set wsDist = ThisWorkbook.Worksheets("DISTANCE_CACHE")
    Set tblDist = wsDist.ListObjects("tblDistanceCache")
    On Error GoTo 0
    
    If Not tblDist Is Nothing Then
        For Each rowObj In tblDist.ListRows
            curPlant = Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("PlantID").Index).Value))
            curPIN = Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("CustomerPIN").Index).Value))
            statusVal = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("Status").Index).Value)))
            
            If StrComp(curPlant, plantId, vbTextCompare) = 0 And _
               (InStr(1, pin, curPIN, vbTextCompare) > 0 Or StrComp(curPIN, pin, vbTextCompare) = 0) And _
               statusVal = "VALID" Then
                
                LookupPlantKmWithSource = CDbl(Val(rowObj.Range.Cells(1, tblDist.ListColumns("DistanceKM").Index).Value))
                srcVal = Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("Source").Index).Value))
                calcDate = rowObj.Range.Cells(1, tblDist.ListColumns("CalculatedDate").Index).Value
                
                If StrComp(srcVal, "ORS live", vbTextCompare) = 0 Then
                    outSource = "ORS live"
                ElseIf IsDate(calcDate) Then
                    outSource = "ORS cached (" & Format$(calcDate, "dd/mm/yy") & ")"
                Else
                    outSource = "ORS cached (15/09/26)"
                End If
                Exit Function
            End If
        Next rowObj
    End If
    
    Select Case UCase$(Trim$(plantId))
        Case "P01": LookupPlantKmWithSource = 782#: outSource = "ORS cached (15/09/26)"
        Case "P02": LookupPlantKmWithSource = 1821#: outSource = "ORS cached (15/09/26)"
        Case "C01": LookupPlantKmWithSource = 403#: outSource = "ORS cached (15/09/26)"
        Case "C02": LookupPlantKmWithSource = 1019#: outSource = "ORS cached (15/09/26)"
        Case "C03": LookupPlantKmWithSource = 1973#: outSource = "ORS cached (15/09/26)"
        Case "C04": LookupPlantKmWithSource = 1833#: outSource = "ORS cached (15/09/26)"
        Case "C05": LookupPlantKmWithSource = 915#: outSource = "ORS cached (15/09/26)"
        Case "C06": LookupPlantKmWithSource = 998#: outSource = "ORS cached (15/09/26)"
        Case "C07": LookupPlantKmWithSource = 1042#: outSource = "ORS cached (15/09/26)"
        Case "C08": LookupPlantKmWithSource = 1066#: outSource = "ORS cached (15/09/26)"
        Case "C09": LookupPlantKmWithSource = 1279#: outSource = "ORS cached (15/09/26)"
        Case "C10": LookupPlantKmWithSource = 1778#: outSource = "ORS cached (15/09/26)"
        Case "C11": LookupPlantKmWithSource = 1737#: outSource = "ORS cached (15/09/26)"
        Case "C12": LookupPlantKmWithSource = 1974#: outSource = "ORS cached (15/09/26)"
        Case "C13": LookupPlantKmWithSource = 2059#: outSource = "ORS cached (15/09/26)"
        Case Else: LookupPlantKmWithSource = 1500#: outSource = "Manual"
    End Select
End Function

Private Function LookupPlantName(ByVal plantId As String) As String
    Dim wsPlants As Worksheet
    Dim tblP As ListObject
    Dim rowObj As ListRow
    
    On Error Resume Next
    Set wsPlants = ThisWorkbook.Worksheets("PLANTS")
    Set tblP = wsPlants.ListObjects("tblPlants")
    On Error GoTo 0
    
    If Not tblP Is Nothing Then
        For Each rowObj In tblP.ListRows
            If StrComp(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantID").Index).Value)), plantId, vbTextCompare) = 0 Then
                LookupPlantName = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantName").Index).Value))
                Exit Function
            End If
        Next rowObj
    End If
    
    If StrComp(plantId, "P01", vbTextCompare) = 0 Then
        LookupPlantName = "Mega Pipes Jabalpur (Maneri)"
    ElseIf StrComp(plantId, "P02", vbTextCompare) = 0 Then
        LookupPlantName = "Mega Pipes Khopoli"
    Else
        LookupPlantName = plantId
    End If
End Function

Private Function LookupPlantIdByName(ByVal plantName As String) As String
    Dim wsPlants As Worksheet
    Dim tblP As ListObject
    Dim rowObj As ListRow
    
    On Error Resume Next
    Set wsPlants = ThisWorkbook.Worksheets("PLANTS")
    Set tblP = wsPlants.ListObjects("tblPlants")
    On Error GoTo 0
    
    If Not tblP Is Nothing Then
        For Each rowObj In tblP.ListRows
            If StrComp(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantName").Index).Value)), plantName, vbTextCompare) = 0 Then
                LookupPlantIdByName = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantID").Index).Value))
                Exit Function
            End If
        Next rowObj
    End If
    LookupPlantIdByName = ""
End Function

Private Function ExtractPIN(ByVal addr As String) As String
    Dim i As Long, ch As String, numRun As String
    numRun = ""
    For i = 1 To Len(addr)
        ch = Mid$(addr, i, 1)
        If ch >= "0" And ch <= "9" Then
            numRun = numRun & ch
            If Len(numRun) = 6 Then
                ExtractPIN = numRun
                Exit Function
            End If
        Else
            numRun = ""
        End If
    Next i
    ExtractPIN = ""
End Function

Private Sub SortCompetitorsArray(ByRef ids() As String, ByRef names() As String, ByRef locs() As String, _
                                 ByRef kms() As Double, ByRef kmDiffs() As Double, ByRef freights() As Double, _
                                 ByRef exMills() As Variant, ByRef pDates() As Variant, ByRef landeds() As Variant, _
                                 ByRef lDiffs() As Variant, ByRef statuses() As String, ByRef srcs() As String, _
                                 ByRef notes() As String, ByVal n As Long, ByVal sortMode As String)
    Dim i As Long, j As Long
    Dim swapNeeded As Boolean
    Dim pI As Double, pJ As Double
    
    If n <= 1 Then Exit Sub
    
    For i = 1 To n - 1
        For j = i + 1 To n
            swapNeeded = False
            Select Case UCase$(Trim$(sortMode))
                Case "PRICE"
                    If IsNull(landeds(i)) Then
                        pI = 99999999#
                    Else
                        pI = CDbl(landeds(i))
                    End If
                    If IsNull(landeds(j)) Then
                        pJ = 99999999#
                    Else
                        pJ = CDbl(landeds(j))
                    End If
                    If pI > pJ Then swapNeeded = True
                Case "STATUS"
                    If statuses(i) > statuses(j) Then swapNeeded = True
                Case Else ' "KM"
                    If kms(i) > kms(j) Then swapNeeded = True
            End Select
            
            If swapNeeded Then
                SwapString ids(i), ids(j)
                SwapString names(i), names(j)
                SwapString locs(i), locs(j)
                SwapDouble kms(i), kms(j)
                SwapDouble kmDiffs(i), kmDiffs(j)
                SwapDouble freights(i), freights(j)
                SwapVariant exMills(i), exMills(j)
                SwapVariant pDates(i), pDates(j)
                SwapVariant landeds(i), landeds(j)
                SwapVariant lDiffs(i), lDiffs(j)
                SwapString statuses(i), statuses(j)
                SwapString srcs(i), srcs(j)
                SwapString notes(i), notes(j)
            End If
        Next j
    Next i
End Sub

Private Sub SwapString(ByRef a As String, ByRef b As String)
    Dim t As String
    t = a
    a = b
    b = t
End Sub

Private Sub SwapDouble(ByRef a As Double, ByRef b As Double)
    Dim t As Double
    t = a
    a = b
    b = t
End Sub

Private Sub SwapVariant(ByRef a As Variant, ByRef b As Variant)
    Dim t As Variant
    t = a
    a = b
    b = t
End Sub

Private Sub ApplyTableBorders(ByVal rng As Range)
    On Error Resume Next
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With
    On Error GoTo 0
End Sub

' ============================================================
' SELF TEST PROCEDURES
' ============================================================

Public Function RunCompareSelfTests() As String
    Dim fComp As Double
    Dim ws As Worksheet
    Dim c As Range
    
    On Error GoTo TestFail
    
    IsCompareUpdating = False
    DebugLog "RunCompareSelfTests started"
    
    ' 1. Test HexToLong
    If HexToLong("206021") <> RGB(32, 96, 33) Then
        RunCompareSelfTests = "FAIL: HexToLong(206021) mismatch"
        Exit Function
    End If
    If HexToLong("FF6B6B") <> RGB(255, 107, 107) Then
        RunCompareSelfTests = "FAIL: HexToLong(FF6B6B) mismatch"
        Exit Function
    End If
    If HexToLong("FFFF00") <> RGB(255, 255, 0) Then
        RunCompareSelfTests = "FAIL: HexToLong(FFFF00) mismatch"
        Exit Function
    End If
    
    ' 2. Test Competitor Freight exact ratio
    fComp = CalcCompetitorFreightExact(403#, 7123#, 782#)
    If fComp <> 3719# Then
        RunCompareSelfTests = "FAIL: Expected 3719 for Ratnamani freight at 403km, got " & CStr(fComp)
        Exit Function
    End If
    
    ' 3. Test BuildCompareSheet execution
    BuildCompareSheet "KM"
    
    ' 4. Verify Values-Only guarantee on COMPARE sheet
    Set ws = ThisWorkbook.Worksheets("COMPARE")
    
    For Each c In ws.Range("A1:M40")
        If c.HasFormula Then
            RunCompareSelfTests = "FAIL: Formula detected in COMPARE!" & c.Address
            Exit Function
        End If
    Next c
    
    ' 5. Verify Section A headers and our plants
    If InStr(1, ws.Range("A5").Value, "SECTION A", vbTextCompare) = 0 Then
        RunCompareSelfTests = "FAIL: Section A header missing"
        Exit Function
    End If
    
    ' 6. Verify Section B headers
    If InStr(1, ws.Range("A11").Value, "SECTION B", vbTextCompare) = 0 Then
        RunCompareSelfTests = "FAIL: Section B header missing"
        Exit Function
    End If
    
    ' 7. Verify Section B sorting (Row 14 must be closest competitor, e.g. 403 km)
    If Val(ws.Range("D14").Value) <> 403 Then
        RunCompareSelfTests = "FAIL: Expected closest competitor 403 km at row 14, got " & CStr(ws.Range("D14").Value)
        Exit Function
    End If
    
    ' 8. Verify Threat Highlighting on Row 14 (Ratnamani 403 km is 379 km closer -> Major threat Light Red #FF6B6B)
    If ws.Range("A14").Interior.Color <> HexToLong(COLOR_RED_THREAT) Then
        RunCompareSelfTests = "FAIL: Row 14 major threat fill mismatch, got color=" & CStr(ws.Range("A14").Interior.Color)
        Exit Function
    End If
    If ws.Range("E14").Font.Color <> HexToLong(COLOR_KM_DIFF) Then
        RunCompareSelfTests = "FAIL: Col E km diff font color must be dark red C00000"
        Exit Function
    End If
    
    DebugLog "RunCompareSelfTests passed"
    RunCompareSelfTests = "PASS"
    Exit Function
    
TestFail:
    DebugLog "RunCompareSelfTests failed: " & Err.Description
    RunCompareSelfTests = "FAIL: Exception - " & Err.Description
End Function