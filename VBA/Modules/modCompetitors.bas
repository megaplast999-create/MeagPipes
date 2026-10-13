Attribute VB_Name = "modCompetitors"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modCompetitors
' Purpose:
'   P10 Competitor engine
'   Prefills prices and dates from tblCompPrices
'   Calculates landed prices and market freight
'   Enforces floor protection for recommended offer
'   Applies runtime yellow highlighting for theirKm < ourNearestKm
' ============================================================

Public Function GetCompetitorPrice(ByVal competitorId As String, ByVal productId As String, _
                                   ByRef outPrice As Variant, ByRef outDate As Variant, ByRef outStatus As String) As Boolean
    Dim wsMasters As Worksheet
    Dim tblCP As ListObject
    Dim rowObj As ListRow
    Dim cId As String
    Dim pId As String
    Dim priceVal As Variant
    Dim dateVal As Variant
    Dim staleDays As Long
    Dim ageDays As Long
    Dim found As Boolean
    
    outPrice = Null
    outDate = Null
    outStatus = "price needed"
    found = False
    
    On Error Resume Next
    Set wsMasters = ThisWorkbook.Worksheets("MASTERS")
    Set tblCP = wsMasters.ListObjects("tblCompPrices")
    On Error GoTo 0
    
    If tblCP Is Nothing Then Exit Function
    
    For Each rowObj In tblCP.ListRows
        cId = Trim$(CStr(rowObj.Range.Cells(1, tblCP.ListColumns("CompetitorID").Index).Value))
        pId = Trim$(CStr(rowObj.Range.Cells(1, tblCP.ListColumns("ProductID").Index).Value))
        
        If StrComp(cId, competitorId, vbTextCompare) = 0 And StrComp(pId, productId, vbTextCompare) = 0 Then
            priceVal = rowObj.Range.Cells(1, tblCP.ListColumns("ExMillRsPerMT").Index).Value
            dateVal = rowObj.Range.Cells(1, tblCP.ListColumns("PriceDate").Index).Value
            found = True
            Exit For
        End If
    Next rowObj
    
    If Not found Or IsEmpty(priceVal) Or IsNull(priceVal) Or Val(priceVal) <= 0 Then
        outStatus = "price needed"
        GetCompetitorPrice = False
        Exit Function
    End If
    
    outPrice = CDec(priceVal)
    outDate = dateVal
    
    ' Check staleness
    On Error Resume Next
    staleDays = CLng(Val(SettingValue("StaleDays")))
    If staleDays <= 0 Then staleDays = 30
    On Error GoTo 0
    
    If IsDate(dateVal) Then
        ageDays = DateDiff("d", CDate(dateVal), Date)
        If ageDays > staleDays Then
            outStatus = "STALE"
        Else
            outStatus = "CURRENT"
        End If
    Else
        outStatus = "STALE"
    End If
    
    GetCompetitorPrice = True
End Function

Public Function CalcCompetitorFreight(ByVal theirKm As Double, ByVal ourFreightPerMT As Double, ByVal ourNearestKm As Double) As Double
    Dim marketRate As Double
    On Error Resume Next
    marketRate = CDbl(Val(SettingValue("MarketFreightRsPerMTperKm")))
    On Error GoTo 0
    
    If marketRate > 0 Then
        CalcCompetitorFreight = CDbl(RoundUpD(CDec(theirKm) * CDec(marketRate), 0))
    Else
        If ourNearestKm > 0 Then
            CalcCompetitorFreight = CDbl(RoundUpD((CDec(ourFreightPerMT) / CDec(ourNearestKm)) * CDec(theirKm), 0))
        Else
            CalcCompetitorFreight = CDbl(RoundUpD(CDec(theirKm) * CDec("10.0"), 0))
        End If
    End If
End Function

Public Function CalcLandedPrice(ByVal exMill As Double, ByVal freight As Double) As Double
    CalcLandedPrice = exMill + freight
End Function

Public Function GetRecommendedOffer(ByVal lowestCompetitorLanded As Double, ByVal floorLanded As Double, _
                                    ByVal undercutAmount As Double, ByRef outWarning As String) As Double
    Dim targetOffer As Double
    targetOffer = lowestCompetitorLanded - undercutAmount
    
    If targetOffer < floorLanded Then
        GetRecommendedOffer = floorLanded
        outWarning = "FLOOR_PROTECTED: Floor price (" & Format$(floorLanded, "0.00") & _
                     ") exceeds competitor target (" & Format$(targetOffer, "0.00") & ")"
    Else
        GetRecommendedOffer = targetOffer
        outWarning = "OK"
    End If
End Function

Public Function IsYellowHighlighted(ByVal theirKm As Double, ByVal ourNearestPlantKm As Double) As Boolean
    ' Runtime yellow highlight rule: their km < our nearest plant km gets FFFF00 yellow background
    IsYellowHighlighted = (theirKm < ourNearestPlantKm)
End Function

Public Sub BuildCompetitorTable(ByVal productId As String, ByVal deliveryPIN As String, _
                                ByVal ourExMillMT As Double, ByVal ourFreightMT As Double, _
                                ByVal ourNearestKm As Double)
    Dim wsComp As Worksheet
    Dim wsPlants As Worksheet
    Dim tblP As ListObject
    Dim rowObj As ListRow
    Dim pType As String
    Dim cId As String
    Dim cName As String
    Dim dist As Double
    Dim exMill As Variant
    Dim pDate As Variant
    Dim pStatus As String
    Dim compFreight As Double
    Dim landed As Double
    Dim diffVsOurs As Double
    Dim ourLanded As Double
    Dim startRow As Long
    Dim rIdx As Long
    Dim isYellow As Boolean
    Dim lowestLanded As Double
    Dim recOffer As Double
    Dim warn As String
    Dim undercut As Double
    
    Set wsComp = ThisWorkbook.Worksheets("COMPARE")
    Set wsPlants = ThisWorkbook.Worksheets("PLANTS")
    Set tblP = wsPlants.ListObjects("tblPlants")
    
    ourLanded = ourExMillMT + ourFreightMT
    lowestLanded = 999999#
    
    On Error Resume Next
    undercut = CDbl(Val(SettingValue("UndercutRsPerMT")))
    If undercut <= 0 Then undercut = 500#
    On Error GoTo 0
    
    ' Clear competitor table area
    wsComp.Range("A13:H25").Clear
    
    wsComp.Range("A12").Value = "COMPETITOR COMPARISON"
    wsComp.Range("A12").Font.Bold = True
    wsComp.Range("A12").Font.Size = 12
    
    ' Headers
    wsComp.Range("A13").Value = "Competitor Name"
    wsComp.Range("B13").Value = "Road KM"
    wsComp.Range("C13").Value = "Ex-Mill / MT"
    wsComp.Range("D13").Value = "Freight / MT"
    wsComp.Range("E13").Value = "Landed / MT"
    wsComp.Range("F13").Value = "Diff vs Ours"
    wsComp.Range("G13").Value = "Price Status"
    wsComp.Range("H13").Value = "Proximity Note"
    
    With wsComp.Range("A13:H13")
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238) ' Light blue
    End With
    
    startRow = 14
    rIdx = 0
    
    For Each rowObj In tblP.ListRows
        pType = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantType").Index).Value)))
        If pType = "COMPETITOR" Then
            cId = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantID").Index).Value))
            cName = Trim$(CStr(rowObj.Range.Cells(1, tblP.ListColumns("PlantName").Index).Value))
            
            ' Distance
            dist = CDbl(Val(CacheLookup("driving-hgv", cId, deliveryPIN)))
            If dist <= 0 Then dist = 1500#
            
            ' Price
            GetCompetitorPrice cId, productId, exMill, pDate, pStatus
            
            If Not IsNull(exMill) Then
                compFreight = CalcCompetitorFreight(dist, ourFreightMT, ourNearestKm)
                landed = CalcLandedPrice(CDbl(exMill), compFreight)
                diffVsOurs = landed - ourLanded
                If landed < lowestLanded Then lowestLanded = landed
            Else
                compFreight = 0
                landed = 0
                diffVsOurs = 0
            End If
            
            isYellow = IsYellowHighlighted(dist, ourNearestKm)
            
            wsComp.Cells(startRow + rIdx, 1).Value = cName
            wsComp.Cells(startRow + rIdx, 2).Value = dist
            wsComp.Cells(startRow + rIdx, 3).Value = IIf(IsNull(exMill), "N/A", exMill)
            wsComp.Cells(startRow + rIdx, 4).Value = compFreight
            wsComp.Cells(startRow + rIdx, 5).Value = landed
            wsComp.Cells(startRow + rIdx, 6).Value = IIf(landed > 0, diffVsOurs, "N/A")
            wsComp.Cells(startRow + rIdx, 7).Value = pStatus
            wsComp.Cells(startRow + rIdx, 8).Value = IIf(isYellow, "CLOSER THAN OUR PLANT", "")
            
            If isYellow Then
                wsComp.Range(wsComp.Cells(startRow + rIdx, 1), wsComp.Cells(startRow + rIdx, 8)).Interior.Color = RGB(255, 255, 0)
            End If
            
            rIdx = rIdx + 1
        End If
    Next rowObj
    
    ' Summary
    If lowestLanded < 999999# Then
        recOffer = GetRecommendedOffer(lowestLanded, ourLanded, undercut, warn)
        wsComp.Cells(startRow + rIdx + 1, 1).Value = "Recommended Offer (Undercut Rs " & CStr(undercut) & "):"
        wsComp.Cells(startRow + rIdx + 1, 2).Value = recOffer
        wsComp.Cells(startRow + rIdx + 1, 4).Value = warn
        wsComp.Range(wsComp.Cells(startRow + rIdx + 1, 1), wsComp.Cells(startRow + rIdx + 1, 4)).Font.Bold = True
    End If
End Sub

Public Function RunCompetitorSelfTests() As String
    Dim pVal As Variant
    Dim pDate As Variant
    Dim pStatus As String
    Dim res As Boolean
    Dim fr As Double
    Dim landed As Double
    Dim rec As Double
    Dim warn As String
    
    On Error GoTo TestFail
    
    ' 1. Test CURRENT price (C01 fresh)
    res = GetCompetitorPrice("C01", "P02", pVal, pDate, pStatus)
    If Not res Or pStatus <> "CURRENT" Or pVal <> CDec(79500) Then
        RunCompetitorSelfTests = "FAIL: Expected CURRENT 79500 for C01, got status=" & pStatus & ", price=" & CStr(pVal)
        Exit Function
    End If
    
    ' 2. Test STALE badge (C02 old tender date)
    res = GetCompetitorPrice("C02", "P02", pVal, pDate, pStatus)
    If Not res Or pStatus <> "STALE" Then
        RunCompetitorSelfTests = "FAIL: Expected STALE for C02, got " & pStatus
        Exit Function
    End If
    
    ' 3. Test MISSING price (C03 no price)
    res = GetCompetitorPrice("C03", "P02", pVal, pDate, pStatus)
    If res Or pStatus <> "price needed" Then
        RunCompetitorSelfTests = "FAIL: Expected 'price needed' for C03, got " & pStatus
        Exit Function
    End If
    
    ' 4. Test Competitor freight ratio fallback
    ' theirKm = 403, ourFreight = 8147, ourKm = 782 -> ratio = 8147/782 = 10.418 -> 403 * 10.418 = 4198.5 -> RoundUpD = 4199
    fr = CalcCompetitorFreight(403#, 8147#, 782#)
    If fr <> 4199# Then
        RunCompetitorSelfTests = "FAIL: Expected freight 4199 for 403 km, got " & CStr(fr)
        Exit Function
    End If
    
    ' 5. Test Landed price: 79500 + 4199 = 83699
    landed = CalcLandedPrice(79500#, fr)
    If landed <> 83699# Then
        RunCompetitorSelfTests = "FAIL: Expected landed 83699, got " & CStr(landed)
        Exit Function
    End If
    
    ' 6. Test Recommended offer - Normal undercutting case
    ' Lowest competitor landed = 90000, floor = 80000, undercut = 500 -> 89500
    rec = GetRecommendedOffer(90000#, 80000#, 500#, warn)
    If rec <> 89500# Or warn <> "OK" Then
        RunCompetitorSelfTests = "FAIL: Normal undercutting failed"
        Exit Function
    End If
    
    ' 7. Test Floor protection: Competitor = 78000, undercut = 500 -> target = 77500. Floor = 80000 -> must return 80000!
    rec = GetRecommendedOffer(78000#, 80000#, 500#, warn)
    If rec <> 80000# Or InStr(1, warn, "FLOOR_PROTECTED") = 0 Then
        RunCompetitorSelfTests = "FAIL: Floor protection failed, rec=" & CStr(rec) & ", warn=" & warn
        Exit Function
    End If
    
    ' 8. Test Yellow highlight rule
    If Not IsYellowHighlighted(403#, 782#) Then
        RunCompetitorSelfTests = "FAIL: Ratnamani 403 < 782 must be yellow"
        Exit Function
    End If
    If IsYellowHighlighted(1019#, 782#) Then
        RunCompetitorSelfTests = "FAIL: Welspun 1019 > 782 must not be yellow"
        Exit Function
    End If
    
    ' 9. Test BuildCompetitorTable
    BuildCompetitorTable "P02", "825322", 80570#, 8147#, 782#
    
    RunCompetitorSelfTests = "PASS"
    Exit Function
    
TestFail:
    RunCompetitorSelfTests = "FAIL: Exception - " & Err.Description
End Function
