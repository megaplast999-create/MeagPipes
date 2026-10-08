Attribute VB_Name = "modValidation"
Option Explicit

Public Function IsValidPipeLength(ByVal pipeLengthM As Double) As Boolean
    IsValidPipeLength = (pipeLengthM >= 6# And pipeLengthM <= 13#)
End Function

Public Function IsValidQuantity(ByVal quantity As Double) As Boolean
    IsValidQuantity = (quantity > 0#)
End Function

Public Function IsValidGeometry(ByVal odMm As Double, ByVal idMm As Double, ByVal thicknessMm As Double) As Boolean
    IsValidGeometry = (odMm > 0# And idMm > 0# And thicknessMm > 0# And odMm > idMm)
End Function
