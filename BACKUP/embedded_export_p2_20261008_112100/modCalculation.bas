Attribute VB_Name = "modCalculation"
Option Explicit

Public Function ConvertQuantityToMetres(ByVal quantity As Double, ByVal quantityUnit As String, ByVal massMtPerM As Double, ByVal pipeLengthM As Double) As Double
    If quantity <= 0# Or massMtPerM <= 0# Then
        Err.Raise vbObjectError + 1101, "ConvertQuantityToMetres", "Quantity and mass must be positive."
    End If
    If pipeLengthM < 6# Or pipeLengthM > 13# Then
        Err.Raise vbObjectError + 1102, "ConvertQuantityToMetres", "Pipe length must be between 6 and 13 metres."
    End If

    Select Case UCase$(Trim$(quantityUnit))
        Case "METRES"
            ConvertQuantityToMetres = quantity
        Case "MT"
            ConvertQuantityToMetres = quantity / massMtPerM
        Case "NOS"
            ConvertQuantityToMetres = quantity * pipeLengthM
        Case Else
            Err.Raise vbObjectError + 1103, "ConvertQuantityToMetres", "Unit must be Metres, MT, or Nos."
    End Select
End Function

Public Function CalculatePipeCount(ByVal metres As Double, ByVal pipeLengthM As Double) As Long
    If metres <= 0# Or pipeLengthM <= 0# Then
        Err.Raise vbObjectError + 1104, "CalculatePipeCount", "Metres and pipe length must be positive."
    End If
    CalculatePipeCount = WorksheetFunction.RoundUp(metres / pipeLengthM, 0)
End Function

Public Function CalculateWeightMT(ByVal metres As Double, ByVal massKgPerM As Double) As Double
    If metres < 0# Or massKgPerM <= 0# Then
        Err.Raise vbObjectError + 1105, "CalculateWeightMT", "Invalid quantity or mass."
    End If
    CalculateWeightMT = metres * massKgPerM / 1000#
End Function

Public Function CalculateCoatingCostPerMT(ByVal rateRsPerM2 As Double, ByVal diameterMm As Double, ByVal massKgPerM As Double) As Double
    If rateRsPerM2 < 0# Or diameterMm <= 0# Or massKgPerM <= 0# Then
        Err.Raise vbObjectError + 1106, "CalculateCoatingCostPerMT", "Invalid coating inputs."
    End If
    CalculateCoatingCostPerMT = rateRsPerM2 * WorksheetFunction.Pi() * diameterMm / 1000# / massKgPerM
End Function

Public Function CalculateFinancialCost(ByVal preCost As Double, ByVal externalCost As Double, ByVal internalCost As Double, ByVal freightCost As Double, ByVal financialPct As Double, ByVal gstPct As Double) As Double
    Dim k As Double
    k = (financialPct / 100#) * (1# + gstPct / 100#)
    If k < 0# Or k >= 1# Then
        Err.Raise vbObjectError + 1107, "CalculateFinancialCost", "Invalid financial/GST factor."
    End If
    CalculateFinancialCost = k * (preCost + externalCost + internalCost + freightCost) / (1# - k)
End Function
