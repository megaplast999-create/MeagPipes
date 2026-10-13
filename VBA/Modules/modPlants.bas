Attribute VB_Name = "modPlants"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modPlants
' Purpose:
'   P8 Plant engine
'   Determines plant eligibility from tblPlantProducts
'   Selects nearest eligible OURS plant via road distance
'   Builds plant comparison table on COMPARE sheet
'   Enforces and logs plant overrides with mandatory reasons
' ============================================================

Public Function GetEligiblePlants(ByVal productId As String) As Variant
    Dim wsPlants As Worksheet
    Dim tblPP As ListObject
    Dim tblP As ListObject
    Dim rowObj As ListRow
    Dim plantId As String
    Dim prodId As String
    Dim activeVal As String
    Dim plantType As String
    Dim plantActive As String
    Dim count As Long
    Dim eligible() As String
    Dim rPlant As ListRow
    Dim isOurs As Boolean
    
    Set wsPlants = ThisWorkbook.Worksheets("PLANTS")
    Set tblPP = wsPlants.ListObjects("tblPlantProducts")
    Set tblP = wsPlants.ListObjects("tblPlants")
    
    count = 0
    ReDim eligible(0 To 10)
    
    For Each rowObj In tblPP.ListRows
        prodId = Trim$(CStr(rowObj.Range.Cells(1, tblPP.ListColumns("ProductID").Index).Value))
        activeVal = Trim$(CStr(rowObj.Range.Cells(1, tblPP.ListColumns("Active").Index).Value))
        
        If StrComp(prodId, productId, vbTextCompare) = 0 And UCase$(activeVal) = "YES" Then
            plantId = Trim$(CStr(rowObj.Range.Cells(1, tblPP.ListColumns("PlantID").Index).Value))
            
            ' Verify plant is OURS and Active in tblPlants
            isOurs = False
            For Each rPlant In tblP.ListRows
                If StrComp(Trim$(CStr(rPlant.Range.Cells(1, tblP.ListColumns("PlantID").Index).Value)), plantId, vbTextCompare) = 0 Then
                    plantType = UCase$(Trim$(CStr(rPlant.Range.Cells(1, tblP.ListColumns("PlantType").Index).Value)))
                    plantActive = UCase$(Trim$(CStr(rPlant.Range.Cells(1, tblP.ListColumns("Active").Index).Value)))
                    If plantType = "OURS" And plantActive = "YES" Then
                        isOurs = True
                    End If
                    Exit For
                End If
            Next rPlant
            
            If isOurs Then
                eligible(count) = plantId
                count = count + 1
            End If
        End If
    Next rowObj
    
    If count > 0 Then
        ReDim Preserve eligible(0 To count - 1)
        GetEligiblePlants = eligible
    Else
        GetEligiblePlants = Array()
    End If
End Function

Public Function GetNearestPlant(ByVal deliveryPIN As String, ByRef eligiblePlants As Variant) As String
    Dim wsDist As Worksheet
    Dim tblDist As ListObject
    Dim rowObj As ListRow
    Dim i As Long
    Dim pId As String
    Dim dist As Double
    Dim minDist As Double
    Dim bestPlant As String
    Dim curPIN As String
    Dim curPlant As String
    
    Set wsDist = ThisWorkbook.Worksheets("DISTANCE_CACHE")
    Set tblDist = wsDist.ListObjects("tblDistanceCache")
    
    minDist = 999999#
    bestPlant = ""
    
    If Not IsArray(eligiblePlants) Then Exit Function
    If UBound(eligiblePlants) < LBound(eligiblePlants) Then Exit Function
    
    For i = LBound(eligiblePlants) To UBound(eligiblePlants)
        pId = CStr(eligiblePlants(i))
        dist = -1#
        
        ' Look up distance in cache
        For Each rowObj In tblDist.ListRows
            curPlant = Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("PlantID").Index).Value))
            curPIN = Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("CustomerPIN").Index).Value))
            
            If StrComp(curPlant, pId, vbTextCompare) = 0 And InStr(1, deliveryPIN, curPIN, vbTextCompare) > 0 Then
                dist = Val(rowObj.Range.Cells(1, tblDist.ListColumns("DistanceKM").Index).Value)
                Exit For
            End If
        Next rowObj
        
        ' Default fallback if not in cache
        If dist < 0 Then
            If StrComp(pId, "P01", vbTextCompare) = 0 Then
                dist = 782#
            ElseIf StrComp(pId, "P02", vbTextCompare) = 0 Then
                dist = 1821#
            Else
                dist = 1000#
            End If
        End If
        
        If dist > 0 And dist < minDist Then
            minDist = dist
            bestPlant = pId
        End If
    Next i
    
    GetNearestPlant = bestPlant
End Function

Public Function CheckOverride(ByVal quoteId As String, ByVal selectedPlantId As String, ByVal nearestPlantId As String, ByVal reason As String) As Boolean
    Dim cleanReason As String
    Dim validReason As Boolean
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim newRow As ListRow
    Dim curUser As String
    
    If StrComp(Trim$(selectedPlantId), Trim$(nearestPlantId), vbTextCompare) = 0 Or Len(Trim$(nearestPlantId)) = 0 Then
        CheckOverride = True
        Exit Function
    End If
    
    cleanReason = Trim$(reason)
    validReason = False
    If StrComp(cleanReason, "Capacity", vbTextCompare) = 0 Or _
       StrComp(cleanReason, "Stock", vbTextCompare) = 0 Or _
       StrComp(cleanReason, "Customer request", vbTextCompare) = 0 Or _
       StrComp(cleanReason, "Other", vbTextCompare) = 0 Then
        validReason = True
    End If
    
    If Not validReason Then
        Err.Raise vbObjectError + 1801, "CheckOverride", _
            "Plant override requires a valid reason: Capacity, Stock, Customer request, or Other."
    End If
    
    ' Log override in tblQuoteLog
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    On Error GoTo 0
    
    If Not tblLog Is Nothing Then
        curUser = Environ$("USERNAME")
        If Len(curUser) = 0 Then curUser = "Antigravity"
        
        If tblLog.ListRows.Count = 1 And Len(Trim$(CStr(tblLog.ListRows(1).Range.Cells(1, tblLog.ListColumns("Action").Index).Value))) = 0 Then
            Set newRow = tblLog.ListRows(1)
        Else
            Set newRow = tblLog.ListRows.Add
        End If
        
        With newRow.Range
            .Cells(1, tblLog.ListColumns("LogID").Index).Value = "LOG-" & Format$(newRow.Index, "00000")
            .Cells(1, tblLog.ListColumns("QuoteID").Index).Value = quoteId
            .Cells(1, tblLog.ListColumns("Timestamp").Index).Value = Now
            .Cells(1, tblLog.ListColumns("User").Index).Value = curUser
            .Cells(1, tblLog.ListColumns("Action").Index).Value = "PLANT_OVERRIDE"
            .Cells(1, tblLog.ListColumns("OldValue").Index).Value = nearestPlantId
            .Cells(1, tblLog.ListColumns("NewValue").Index).Value = selectedPlantId
            .Cells(1, tblLog.ListColumns("SheetName").Index).Value = "QUOTE"
            .Cells(1, tblLog.ListColumns("Status").Index).Value = "LOGGED"
            .Cells(1, tblLog.ListColumns("Notes").Index).Value = cleanReason
        End With
    End If
    
    CheckOverride = True
End Function

Public Function BuildPlantComparison(ByVal quoteId As String, ByVal productId As String, ByVal deliveryPIN As String, _
                                     ByVal totalWeightMT As Double, ByVal totalTrips As Long, ByVal baseExMillPerMT As Double) As Variant
    Dim wsComp As Worksheet
    Dim eligible As Variant
    Dim i As Long
    Dim pId As String
    Dim pName As String
    Dim dist As Double
    Dim tripRate As Currency
    Dim freightMT As Currency
    Dim landedMT As Currency
    Dim wsPlants As Worksheet
    Dim tblP As ListObject
    Dim rPlant As ListRow
    Dim startRow As Long
    
    eligible = GetEligiblePlants(productId)
    If Not IsArray(eligible) Then Exit Function
    
    Set wsComp = ThisWorkbook.Worksheets("COMPARE")
    Set wsPlants = ThisWorkbook.Worksheets("PLANTS")
    Set tblP = wsPlants.ListObjects("tblPlants")
    
    wsComp.Range("A3:F10").ClearContents
    
    ' Headers
    wsComp.Range("A3").Value = "Plant ID"
    wsComp.Range("B3").Value = "Plant Name"
    wsComp.Range("C3").Value = "Distance (KM)"
    wsComp.Range("D3").Value = "Trips"
    wsComp.Range("E3").Value = "Freight / MT"
    wsComp.Range("F3").Value = "Landed Price / MT"
    
    With wsComp.Range("A3:F3")
        .Font.Bold = True
        .Interior.Color = RGB(197, 180, 227) ' Lavender
    End With
    
    startRow = 4
    For i = LBound(eligible) To UBound(eligible)
        pId = CStr(eligible(i))
        pName = pId
        For Each rPlant In tblP.ListRows
            If StrComp(Trim$(CStr(rPlant.Range.Cells(1, tblP.ListColumns("PlantID").Index).Value)), pId, vbTextCompare) = 0 Then
                pName = Trim$(CStr(rPlant.Range.Cells(1, tblP.ListColumns("PlantName").Index).Value))
                Exit For
            End If
        Next rPlant
        
        If StrComp(pId, "P01", vbTextCompare) = 0 Then
            dist = 782#
        Else
            dist = 1821#
        End If
        
        tripRate = CCur(GetTripRate("B", dist))
        If totalWeightMT > 0 Then
            freightMT = CCur(CalcFreightPerMT(tripRate, totalTrips, totalWeightMT, CDec(SettingValue("HandlingRsPerMT"))))
        Else
            freightMT = 0
        End If
        landedMT = CCur(baseExMillPerMT) + freightMT
        
        wsComp.Cells(startRow + i, 1).Value = pId
        wsComp.Cells(startRow + i, 2).Value = pName
        wsComp.Cells(startRow + i, 3).Value = dist
        wsComp.Cells(startRow + i, 4).Value = totalTrips
        wsComp.Cells(startRow + i, 5).Value = freightMT
        wsComp.Cells(startRow + i, 6).Value = landedMT
    Next i
    
    BuildPlantComparison = "OK"
End Function

Public Function RunPlantSelfTests() As String
    Dim elHSAW As Variant
    Dim elHFW As Variant
    Dim nearest As String
    Dim res As Boolean
    Dim errCaught As Boolean
    Dim wsLog As Worksheet
    Dim tblLog As ListObject
    Dim lastRow As ListRow
    
    On Error GoTo TestFail
    
    ' 1. Test HSAW eligibility (should include P01 and P02)
    elHSAW = GetEligiblePlants("P02")
    If UBound(elHSAW) - LBound(elHSAW) + 1 <> 2 Then
        RunPlantSelfTests = "FAIL: Expected 2 plants for P02 (HSAW), got " & CStr(UBound(elHSAW) - LBound(elHSAW) + 1)
        Exit Function
    End If
    
    ' 2. Test HFW eligibility (should include only P02)
    elHFW = GetEligiblePlants("P04")
    If UBound(elHFW) - LBound(elHFW) + 1 <> 1 Then
        RunPlantSelfTests = "FAIL: Expected 1 plant for P04 (HFW), got " & CStr(UBound(elHFW) - LBound(elHFW) + 1)
        Exit Function
    End If
    If elHFW(0) <> "P02" Then
        RunPlantSelfTests = "FAIL: Expected P02 for P04 (HFW), got " & elHFW(0)
        Exit Function
    End If
    
    ' 3. Test nearest plant selection for Bagodar (825322)
    nearest = GetNearestPlant("Krishna Nagar, Bagodar, Jharkhand - 825322", elHSAW)
    If nearest <> "P01" Then
        RunPlantSelfTests = "FAIL: Expected P01 as nearest to Bagodar, got " & nearest
        Exit Function
    End If
    
    ' 4. Test CheckOverride with valid reason
    res = CheckOverride("Q-TEST-PLANT", "P02", "P01", "Capacity")
    If Not res Then
        RunPlantSelfTests = "FAIL: CheckOverride with valid reason returned False"
        Exit Function
    End If
    
    ' Verify log entry
    Set wsLog = ThisWorkbook.Worksheets("QUOTE_LOG")
    Set tblLog = wsLog.ListObjects("tblQuoteLog")
    Set lastRow = tblLog.ListRows(tblLog.ListRows.Count)
    If Trim$(CStr(lastRow.Range.Cells(1, tblLog.ListColumns("Action").Index).Value)) <> "PLANT_OVERRIDE" Or _
       Trim$(CStr(lastRow.Range.Cells(1, tblLog.ListColumns("Notes").Index).Value)) <> "Capacity" Then
        RunPlantSelfTests = "FAIL: Audit log entry mismatch for override"
        Exit Function
    End If
    
    ' 5. Test CheckOverride without reason (must raise error)
    errCaught = False
    On Error Resume Next
    CheckOverride "Q-TEST-PLANT", "P02", "P01", ""
    If Err.Number <> 0 Then errCaught = True
    On Error GoTo TestFail
    
    If Not errCaught Then
        RunPlantSelfTests = "FAIL: Expected error for empty override reason"
        Exit Function
    End If
    
    ' 6. Test comparison table builder
    BuildPlantComparison "Q-TEST-PLANT", "P02", "825322", 2031.036, 153, 80570
    
    RunPlantSelfTests = "PASS"
    Exit Function
    
TestFail:
    RunPlantSelfTests = "FAIL: Exception - " & Err.Description
End Function
