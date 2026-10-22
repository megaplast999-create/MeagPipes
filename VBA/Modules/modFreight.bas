Attribute VB_Name = "modFreight"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modFreight
' Purpose:
'   P7 Freight engine
'   Supports methods A, B, and C
'   Stacking lookup, trailers, trips, and freight per MT
' ============================================================

Public Function TrailersForLine(ByVal pipes As Variant, ByVal pipesPerTrailer As Variant, Optional ByVal legacyMode As Boolean = False) As Variant
    On Error GoTo ErrorHandler
    If CDec(pipes) <= 0 Or CDec(pipesPerTrailer) <= 0 Then
        TrailersForLine = CDec(0)
        Exit Function
    End If
    
    If legacyMode Then
        ' Fractional trailers in legacy mode
        TrailersForLine = CDec(pipes) / CDec(pipesPerTrailer)
    Else
        ' Whole trailers rounded up in normal mode
        TrailersForLine = RoundUpD(CDec(pipes) / CDec(pipesPerTrailer), 0)
    End If
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, "TrailersForLine", Err.Description
End Function

Public Function GetPipesPerTrailer(ByVal odMM As Double, ByVal pipeWeightMT As Double) As Long
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowObj As ListRow
    Dim minOD As Double, maxOD As Double
    Dim found As Boolean
    Dim payloadMT As Double
    
    found = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("MASTERS")
    If Not ws Is Nothing Then
        Set tbl = ws.ListObjects("tblStacking")
    End If
    On Error GoTo 0
    
    If Not tbl Is Nothing Then
        For Each rowObj In tbl.ListRows
            minOD = Val(rowObj.Range.Cells(1, tbl.ListColumns("MinOD_MM").Index).Value)
            maxOD = Val(rowObj.Range.Cells(1, tbl.ListColumns("MaxOD_MM").Index).Value)
            If odMM >= minOD And odMM <= maxOD Then
                GetPipesPerTrailer = CLng(Val(rowObj.Range.Cells(1, tbl.ListColumns("PipesPerTrailer").Index).Value))
                found = True
                Exit For
            End If
        Next rowObj
    End If
    
    If Not found Then
        On Error Resume Next
        payloadMT = CDec(SettingValue("TrailerPayloadMT"))
        If Err.Number <> 0 Or payloadMT <= 0 Then payloadMT = 28#
        On Error GoTo 0
        
        If pipeWeightMT > 0 Then
            GetPipesPerTrailer = Fix(payloadMT / pipeWeightMT)
        Else
            GetPipesPerTrailer = 1
        End If
    End If
    
    If GetPipesPerTrailer <= 0 Then GetPipesPerTrailer = 1
End Function

Public Function GetTripRate(ByVal method As String, ByVal km As Double, Optional ByVal manualRate As Variant) As Variant
    Dim cleanMethod As String
    cleanMethod = UCase$(Trim$(method))
    
    Select Case cleanMethod
        Case "A"
            ' Method A: Manual trip rate
            If Not IsMissing(manualRate) Then
                If CDec(manualRate) > 0 Then
                    GetTripRate = CDec(manualRate)
                    Exit Function
                End If
            End If
            GetTripRate = CDec(SettingValue("TripRateManual"))
            
        Case "B"
            ' Method B: RatePerKm * km
            GetTripRate = CDec(SettingValue("RatePerKm")) * CDec(km)
            
        Case "C"
            ' Method C: Slab-based rate from tblSlabs
            Dim ws As Worksheet
            Dim tbl As ListObject
            Dim rowObj As ListRow
            Dim minKm As Double, maxKm As Double
            Dim found As Boolean
            found = False
            
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets("MASTERS")
            If Not ws Is Nothing Then
                Set tbl = ws.ListObjects("tblSlabs")
            End If
            On Error GoTo 0
            
            If Not tbl Is Nothing Then
                For Each rowObj In tbl.ListRows
                    minKm = Val(rowObj.Range.Cells(1, tbl.ListColumns("MinKM").Index).Value)
                    maxKm = Val(rowObj.Range.Cells(1, tbl.ListColumns("MaxKM").Index).Value)
                    If km >= minKm And km <= maxKm Then
                        GetTripRate = CDec(rowObj.Range.Cells(1, tbl.ListColumns("TripRate").Index).Value)
                        found = True
                        Exit For
                    End If
                Next rowObj
            End If
            
            If Not found Then
                ' Fallback to Method B
                GetTripRate = CDec(SettingValue("RatePerKm")) * CDec(km)
            End If
            
        Case Else
            Err.Raise vbObjectError + 1720, "GetTripRate", "Unknown freight method: " & method
    End Select
End Function

Public Function CalcFreightPerMT(ByVal tripRate As Variant, ByVal trailers As Variant, ByVal weightMT As Variant, ByVal handling As Variant) As Variant
    If CDec(weightMT) <= 0 Then
        Err.Raise vbObjectError + 1721, "CalcFreightPerMT", "Weight MT must be positive."
    End If
    CalcFreightPerMT = RoundUpD(CDec(tripRate) * CDec(trailers) / CDec(weightMT), 0) + CDec(handling)
End Function

Public Function RunFreightSelfTests() As String
    Dim res As String
    Dim tCount As Variant
    Dim tRate As Variant
    Dim fMT As Variant
    Dim pPerT As Long
    
    On Error GoTo TestFail
    
    ' 1. Test Methods A, B, C
    Dim expA As Variant
    expA = CDec(SettingValue("TripRateManual"))
    tRate = GetTripRate("A", 782)
    If tRate <> expA Then
        RunFreightSelfTests = "FAIL: Method A expected " & CStr(expA) & ", got " & CStr(tRate)
        Exit Function
    End If
    
    tRate = GetTripRate("B", 782)
    If Abs(CDec(tRate) - CDec(98000.24)) > CDec("0.01") Then
        RunFreightSelfTests = "FAIL: Method B expected 98000.24, got " & CStr(tRate)
        Exit Function
    End If
    
    tRate = GetTripRate("C", 782)
    If tRate <> CDec(98000) Then
        RunFreightSelfTests = "FAIL: Method C expected 98000, got " & CStr(tRate)
        Exit Function
    End If
    
    ' 2. Test Stacking lookup
    pPerT = GetPipesPerTrailer(1243, 3.069)
    If pPerT <> 4 Then
        RunFreightSelfTests = "FAIL: Stacking OD 1243 expected 4, got " & CStr(pPerT)
        Exit Function
    End If
    
    pPerT = GetPipesPerTrailer(1417, 3.502)
    If pPerT <> 4 Then
        RunFreightSelfTests = "FAIL: Stacking OD 1417 expected 4, got " & CStr(pPerT)
        Exit Function
    End If
    
    ' 3. Test Trailers whole vs fractional
    tCount = TrailersForLine(167, 4, False)
    If tCount <> CDec(42) Then
        RunFreightSelfTests = "FAIL: Normal trailers expected 42, got " & CStr(tCount)
        Exit Function
    End If
    
    tCount = TrailersForLine(167, 4, True)
    If Abs(CDec(tCount) - CDec("41.75")) > CDec("0.0001") Then
        RunFreightSelfTests = "FAIL: Legacy trailers expected 41.75, got " & CStr(tCount)
        Exit Function
    End If
    
    ' 4. Test Freight per MT Normal vs Legacy
    ' Normal Line 1: 42 trailers, 98000.24 rate, 511.51107672 MT, 100 handling -> 8147
    fMT = CalcFreightPerMT(CDec(98000.24), CDec(42), CDec("511.51107672"), CDec(100))
    If fMT <> CDec(8147) Then
        RunFreightSelfTests = "FAIL: Normal Line 1 freight expected 8147, got " & CStr(fMT)
        Exit Function
    End If
    
    ' Legacy Line 1: 41.66666667 trailers, 98000 rate, 511.51107672 MT, 100 handling -> 8083
    fMT = CalcFreightPerMT(CDec(98000), CDec("41.66666667"), CDec("511.51107672"), CDec(100))
    If fMT <> CDec(8083) Then
        RunFreightSelfTests = "FAIL: Legacy Line 1 freight expected 8083, got " & CStr(fMT)
        Exit Function
    End If
    
    ' 5. Boundary case: zero or negative pipes
    tCount = TrailersForLine(0, 4, False)
    If tCount <> CDec(0) Then
        RunFreightSelfTests = "FAIL: Zero pipes expected 0 trailers"
        Exit Function
    End If
    
    RunFreightSelfTests = "PASS"
    Exit Function
    
TestFail:
    RunFreightSelfTests = "FAIL: Exception - " & Err.Description
End Function
