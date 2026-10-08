Attribute VB_Name = "modGeometry"
Option Explicit

Public Function CalculateOD(ByVal sizeMm As Double, ByVal thicknessMm As Double, ByVal basis As String) As Double
    If sizeMm <= 0# Or thicknessMm <= 0# Then
        Err.Raise vbObjectError + 1001, "CalculateOD", "Size and thickness must be positive."
    End If

    Select Case UCase$(Trim$(basis))
        Case "OD"
            CalculateOD = sizeMm
        Case "ID"
            CalculateOD = WorksheetFunction.RoundUp(sizeMm + (2# * thicknessMm), 0)
        Case Else
            Err.Raise vbObjectError + 1002, "CalculateOD", "Basis must be OD or ID."
    End Select
End Function

Public Function CalculateID(ByVal sizeMm As Double, ByVal thicknessMm As Double, ByVal basis As String) As Double
    Dim odMm As Double

    If sizeMm <= 0# Or thicknessMm <= 0# Then
        Err.Raise vbObjectError + 1003, "CalculateID", "Size and thickness must be positive."
    End If

    If UCase$(Trim$(basis)) = "ID" Then
        CalculateID = sizeMm
    ElseIf UCase$(Trim$(basis)) = "OD" Then
        odMm = sizeMm
        CalculateID = odMm - (2# * thicknessMm)
    Else
        Err.Raise vbObjectError + 1004, "CalculateID", "Basis must be OD or ID."
    End If

    If CalculateID <= 0# Then
        Err.Raise vbObjectError + 1005, "CalculateID", "Calculated ID must be positive."
    End If
End Function

Public Function CalculateMassKgPerM(ByVal odMm As Double, ByVal thicknessMm As Double) As Double
    If odMm <= 0# Or thicknessMm <= 0# Or odMm <= (2# * thicknessMm) Then
        Err.Raise vbObjectError + 1006, "CalculateMassKgPerM", "Invalid geometry."
    End If
    CalculateMassKgPerM = (odMm - thicknessMm) * thicknessMm * 0.0246615
End Function
