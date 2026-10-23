Attribute VB_Name = "modCostSheetParity"
Option Explicit

' ============================================================
' One-click check that the VBA engine equals the reference cost sheet
' ("CS Stacking 10-09-26", Unipro 594) and the Python engine
' (TESTS/test_costsheet_automation.py). Run: RunCostSheetParityTests
' Expected values are transcribed from the reference workbook.
' ============================================================

Private Function RefLine(ByVal idx As Long, ByVal typedFC As Boolean) As LineResult
    Dim ids As Variant, qty As Variant, fcs As Variant
    Dim rates As RateInput
    Dim ld As QuoteLineInput
    ids = Array(1226, 1300, 1350, 1400)
    qty = Array(2000, 350, 500, 4400)
    fcs = Array(1024, 1014, 1011, 1008)

    LoadDefaultRateInput rates, False
    rates.Option1Mode = False
    rates.WholeTrailers = False      ' fractional trips (reference)
    rates.IntCoatOnID = False        ' OD factor (reference)
    rates.FCAlgebraic = True         ' self-consistent FC
    rates.TripsOverride = Empty
    rates.Steel = CDec(62000): rates.YieldPct = CDec(2): rates.Conversion = CDec(4000)
    rates.FinPct = CDec(1)
    rates.ExtRate = CDec(500): rates.IntRate = CDec(250): rates.PI = CDec(3.14)
    rates.TripRate = CDec(93000)
    rates.Handling = CDec(100)
    If typedFC Then rates.ManualFC = CDec(fcs(idx)) Else rates.ManualFC = CDec(0)

    With ld
        .Basis = "ID": .SizeMM = CDec(ids(idx)): .WT = CDec(8.4)
        .Qty = CDec(qty(idx)): .Unit = "Metres": .PipeLen = CDec(12)
        .PerTrailer = CDec(4): .HasExt = True: .HasInt = True
    End With
    RefLine = CalcLine(ld, rates, False)
End Function

Public Function RunCostSheetParityTests() As String
    Dim i As Long
    Dim r As LineResult
    Dim total As Variant
    Dim expFreight As Variant, expM As Variant, expAuto As Variant

    On Error GoTo Fail
    expFreight = Array(7676, 7248, 6985, 6740)
    expM = Array(22351, 23570, 24396, 25220)
    expAuto = Array(22353, 23573, 24400, 25224)

    total = CDec(0)
    For i = 0 To 3
        r = RefLine(i, True)
        If r.FreightPerMT <> expFreight(i) Then RunCostSheetParityTests = "FAIL: freight L" & (i + 1) & " got " & CStr(r.FreightPerMT): Exit Function
        If r.TotalPerM <> expM(i) Then RunCostSheetParityTests = "FAIL: typed-FC price/m L" & (i + 1) & " got " & CStr(r.TotalPerM): Exit Function
        total = total + r.LineValue
    Next i
    If total <> CDec(176117500) Then RunCostSheetParityTests = "FAIL: typed-FC total got " & CStr(total): Exit Function

    For i = 0 To 3
        r = RefLine(i, False)
        If r.TotalPerM <> expAuto(i) Then RunCostSheetParityTests = "FAIL: auto-FC price/m L" & (i + 1) & " got " & CStr(r.TotalPerM): Exit Function
    Next i
    RunCostSheetParityTests = "PASS"
    Exit Function
Fail:
    RunCostSheetParityTests = "FAIL: Exception - " & Err.Description
End Function
