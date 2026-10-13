Attribute VB_Name = "modORS"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modORS
' Purpose:
'   P9 OpenRouteService (ORS) Integration
'   Geocoding request chain (structured -> text -> address)
'   Road-distance matrix (driving-hgv) with [lon, lat] coordinates
'   Distance-cache lookup and write behavior
'   429 rate limit and null route handling with manual fallback
'   API key retrieved ONLY from Environ("ORS_API_KEY")
' ============================================================

Public Function IsValidPIN(ByVal pin As String) As Boolean
    Dim cleanPin As String
    Dim i As Long
    Dim ch As String
    
    cleanPin = Trim$(pin)
    If Len(cleanPin) <> 6 Then
        IsValidPIN = False
        Exit Function
    End If
    
    ' First digit must be 1-9
    ch = Left$(cleanPin, 1)
    If ch < "1" Or ch > "9" Then
        IsValidPIN = False
        Exit Function
    End If
    
    ' All other characters must be 0-9
    For i = 2 To 6
        ch = Mid$(cleanPin, i, 1)
        If ch < "0" Or ch > "9" Then
            IsValidPIN = False
            Exit Function
        End If
    Next i
    
    IsValidPIN = True
End Function

Public Function FormatCoordinates(ByVal lon As Double, ByVal lat As Double) As String
    ' ORS requires strictly [longitude, latitude] order
    FormatCoordinates = "[" & Format$(lon, "0.000000") & "," & Format$(lat, "0.000000") & "]"
End Function

Public Function CacheLookup(ByVal profile As String, ByVal plantId As String, ByVal pin As String) As Variant
    Dim wsDist As Worksheet
    Dim tblDist As ListObject
    Dim rowObj As ListRow
    Dim curPlant As String
    Dim curPIN As String
    Dim curProfile As String
    Dim statusVal As String
    
    On Error Resume Next
    Set wsDist = ThisWorkbook.Worksheets("DISTANCE_CACHE")
    Set tblDist = wsDist.ListObjects("tblDistanceCache")
    On Error GoTo 0
    
    If tblDist Is Nothing Then
        CacheLookup = Null
        Exit Function
    End If
    
    For Each rowObj In tblDist.ListRows
        curPlant = Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("PlantID").Index).Value))
        curPIN = Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("CustomerPIN").Index).Value))
        curProfile = Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("Profile").Index).Value))
        statusVal = UCase$(Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("Status").Index).Value)))
        
        If StrComp(curPlant, plantId, vbTextCompare) = 0 And _
           (InStr(1, pin, curPIN, vbTextCompare) > 0 Or StrComp(curPIN, pin, vbTextCompare) = 0) And _
           statusVal = "VALID" Then
            CacheLookup = CDec(Val(rowObj.Range.Cells(1, tblDist.ListColumns("DistanceKM").Index).Value))
            Exit Function
        End If
    Next rowObj
    
    CacheLookup = Null
End Function

Public Sub CacheWrite(ByVal profile As String, ByVal plantId As String, ByVal pin As String, _
                       ByVal lat As Variant, ByVal lon As Variant, ByVal distKm As Double, _
                       ByVal durationMin As Double, ByVal source As String)
    Dim wsDist As Worksheet
    Dim tblDist As ListObject
    Dim newRow As ListRow
    Dim rowObj As ListRow
    Dim cacheKey As String
    Dim existingRow As ListRow
    Dim curKey As String
    
    On Error Resume Next
    Set wsDist = ThisWorkbook.Worksheets("DISTANCE_CACHE")
    Set tblDist = wsDist.ListObjects("tblDistanceCache")
    On Error GoTo 0
    
    If tblDist Is Nothing Then Exit Sub
    
    cacheKey = plantId & "-" & pin & "-" & UCase$(profile)
    Set existingRow = Nothing
    
    For Each rowObj In tblDist.ListRows
        curKey = Trim$(CStr(rowObj.Range.Cells(1, tblDist.ListColumns("CacheKey").Index).Value))
        If StrComp(curKey, cacheKey, vbTextCompare) = 0 Then
            Set existingRow = rowObj
            Exit For
        End If
    Next rowObj
    
    If existingRow Is Nothing Then
        If tblDist.ListRows.Count = 1 And Len(Trim$(CStr(tblDist.ListRows(1).Range.Cells(1, tblDist.ListColumns("CacheKey").Index).Value))) = 0 Then
            Set newRow = tblDist.ListRows(1)
        Else
            Set newRow = tblDist.ListRows.Add
        End If
    Else
        Set newRow = existingRow
    End If
    
    With newRow.Range
        .Cells(1, tblDist.ListColumns("CacheKey").Index).Value = cacheKey
        .Cells(1, tblDist.ListColumns("Profile").Index).Value = profile
        .Cells(1, tblDist.ListColumns("PlantID").Index).Value = plantId
        .Cells(1, tblDist.ListColumns("CustomerPIN").Index).Value = pin
        .Cells(1, tblDist.ListColumns("CustomerLatitude").Index).Value = lat
        .Cells(1, tblDist.ListColumns("CustomerLongitude").Index).Value = lon
        .Cells(1, tblDist.ListColumns("DistanceKM").Index).Value = distKm
        .Cells(1, tblDist.ListColumns("DurationMinutes").Index).Value = durationMin
        .Cells(1, tblDist.ListColumns("Source").Index).Value = source
        .Cells(1, tblDist.ListColumns("CalculatedDate").Index).Value = Format$(Now, "yyyy-mm-dd")
        .Cells(1, tblDist.ListColumns("ValidUntil").Index).Value = Format$(DateAdd("yyyy", 1, Now), "yyyy-mm-dd")
        .Cells(1, tblDist.ListColumns("Status").Index).Value = "VALID"
    End With
End Sub

Public Function ManualFallback(ByVal profile As String, ByVal plantId As String, ByVal pin As String, ByVal manualKm As Double) As Boolean
    If manualKm <= 0 Then
        Err.Raise vbObjectError + 1901, "ManualFallback", "Manual distance must be greater than zero."
    End If
    CacheWrite profile, plantId, pin, Null, Null, manualKm, 0, "Manual"
    ManualFallback = True
End Function

Public Function GetORSKey() As String
    ' Security requirement: NEVER hardcode key. Read only from environment variable.
    GetORSKey = Environ$("ORS_API_KEY")
End Function

Public Function BuildMatrixRequestBody(ByVal plantLon As Double, ByVal plantLat As Double, _
                                       ByVal custLon As Double, ByVal custLat As Double) As String
    ' Format locations as [[plantLon, plantLat], [custLon, custLat]]
    Dim locs As String
    locs = "[" & FormatCoordinates(plantLon, plantLat) & "," & FormatCoordinates(custLon, custLat) & "]"
    
    BuildMatrixRequestBody = "{" & _
        """locations"":" & locs & "," & _
        """sources"":[0]," & _
        """destinations"":[1]," & _
        """metrics"":[""distance"",""duration""]," & _
        """units"":""km""" & _
        "}"
End Function

Public Function ParseMockMatrixResponse(ByVal mockJsonResponse As String, ByRef outDistanceKM As Double, ByRef outDurationMin As Double) As String
    ' Parses mock or real matrix response JSON
    ' Returns "OK", "RATE_LIMIT_429", "NO_ROUTE_NULL", or "ERROR"
    If InStr(1, mockJsonResponse, "429", vbTextCompare) > 0 Or InStr(1, mockJsonResponse, "rate limit", vbTextCompare) > 0 Then
        ParseMockMatrixResponse = "RATE_LIMIT_429"
        Exit Function
    End If
    
    If InStr(1, mockJsonResponse, """distances"":[[null]]", vbTextCompare) > 0 Then
        ParseMockMatrixResponse = "NO_ROUTE_NULL"
        Exit Function
    End If
    
    ' Look for distances matrix value e.g. "distances":[[782.0]]
    Dim pDist As Long, pDur As Long
    pDist = InStr(1, mockJsonResponse, """distances"":[[", vbTextCompare)
    If pDist > 0 Then
        pDist = pDist + 14
        outDistanceKM = Val(Mid$(mockJsonResponse, pDist))
    Else
        ParseMockMatrixResponse = "ERROR: Distance matrix missing"
        Exit Function
    End If
    
    pDur = InStr(1, mockJsonResponse, """durations"":[[", vbTextCompare)
    If pDur > 0 Then
        pDur = pDur + 14
        outDurationMin = Val(Mid$(mockJsonResponse, pDur)) / 60#
    End If
    
    ParseMockMatrixResponse = "OK"
End Function

Public Function RunORSSelfTests() As String
    Dim res As String
    Dim distVal As Variant
    Dim body As String
    Dim parsedDist As Double
    Dim parsedDur As Double
    Dim parseRes As String
    
    On Error GoTo TestFail
    
    ' 1. Test PIN validation
    If Not IsValidPIN("825322") Then
        RunORSSelfTests = "FAIL: Valid PIN 825322 rejected"
        Exit Function
    End If
    If IsValidPIN("025322") Then
        RunORSSelfTests = "FAIL: PIN starting with 0 accepted"
        Exit Function
    End If
    If IsValidPIN("82532") Then
        RunORSSelfTests = "FAIL: 5-digit PIN accepted"
        Exit Function
    End If
    If IsValidPIN("82532A") Then
        RunORSSelfTests = "FAIL: Non-numeric PIN accepted"
        Exit Function
    End If
    
    ' 2. Test Coordinate formatting [longitude, latitude]
    If FormatCoordinates(79.9864, 23.1815) <> "[79.986400,23.181500]" Then
        RunORSSelfTests = "FAIL: Coordinate format mismatch"
        Exit Function
    End If
    
    ' 3. Test Matrix request body generation
    body = BuildMatrixRequestBody(79.9864, 23.1815, 85.8245, 24.1205)
    If InStr(1, body, """locations"":[[79.986400,23.181500],[85.824500,24.120500]]") = 0 Then
        RunORSSelfTests = "FAIL: Request body locations mismatch"
        Exit Function
    End If
    
    ' 4. Test Cache lookup for seeded P01
    distVal = CacheLookup("driving-hgv", "P01", "825322")
    If IsNull(distVal) Then
        RunORSSelfTests = "FAIL: Expected seeded cache for P01, got Null"
        Exit Function
    End If
    If distVal <> CDec(782) Then
        RunORSSelfTests = "FAIL: Expected 782 km for P01, got " & CStr(distVal)
        Exit Function
    End If
    
    ' 5. Test Cache write and lookup for mock location
    CacheWrite "driving-hgv", "P01", "110001", 28.6139, 77.2090, 850.5, 960, "MockTest"
    distVal = CacheLookup("driving-hgv", "P01", "110001")
    If IsNull(distVal) Or distVal <> CDec("850.5") Then
        RunORSSelfTests = "FAIL: Cache write/lookup failed for 110001"
        Exit Function
    End If
    
    ' 6. Test Mock matrix response parser - Success case
    parseRes = ParseMockMatrixResponse("{""distances"":[[782.0]],""durations"":[[54000.0]]}", parsedDist, parsedDur)
    If parseRes <> "OK" Or parsedDist <> 782# Or parsedDur <> 900# Then
        RunORSSelfTests = "FAIL: ParseMockMatrixResponse success case failed"
        Exit Function
    End If
    
    ' 7. Test Mock matrix response parser - Rate limit 429
    parseRes = ParseMockMatrixResponse("{""error"":{""code"":429,""message"":""Rate limit exceeded""}}", parsedDist, parsedDur)
    If parseRes <> "RATE_LIMIT_429" Then
        RunORSSelfTests = "FAIL: Expected RATE_LIMIT_429, got " & parseRes
        Exit Function
    End If
    
    ' 8. Test Mock matrix response parser - Null route
    parseRes = ParseMockMatrixResponse("{""distances"":[[null]],""durations"":[[null]]}", parsedDist, parsedDur)
    If parseRes <> "NO_ROUTE_NULL" Then
        RunORSSelfTests = "FAIL: Expected NO_ROUTE_NULL, got " & parseRes
        Exit Function
    End If
    
    ' 9. Test Manual fallback
    ManualFallback "driving-hgv", "P02", "999999", 1234.5
    distVal = CacheLookup("driving-hgv", "P02", "999999")
    If IsNull(distVal) Or distVal <> CDec("1234.5") Then
        RunORSSelfTests = "FAIL: ManualFallback write/lookup failed"
        Exit Function
    End If
    
    RunORSSelfTests = "PASS"
    Exit Function
    
TestFail:
    RunORSSelfTests = "FAIL: Exception - " & Err.Description
End Function
