Attribute VB_Name = "modPricing"
Option Explicit

' ============================================================
' Mega Pipes Quotation System
' Module: modPricing
' Purpose:
'   P3 calculation and pricing engine
'
' Important:
'   - UDT parameters are passed ByRef.
'   - QuoteLineInput is the canonical quote-line input type.
'   - RateInput contains all configurable calculation inputs.
'   - LineResult contains calculated outputs.
' ============================================================


' ============================================================
' QUOTE LINE INPUT
' ============================================================

Public Type QuoteLineInput
    Basis As String
    SizeMM As Variant
    WT As Variant
    Qty As Variant
    Unit As String
    PipeLen As Variant
    PerTrailer As Variant
    HasExt As Boolean
    HasInt As Boolean
End Type


' ============================================================
' RATE / COST INPUT
' ============================================================

Public Type RateInput
    Steel As Variant
    GradeExtra As Variant
    YieldPct As Variant
    Conversion As Variant
    Testing As Variant
    Margin As Variant
    FinPct As Variant
    GSTPct As Variant
    PI As Variant
    ExtRate As Variant
    IntRate As Variant
    Handling As Variant
    TripRate As Variant
    ManualFC As Variant
    ManualSellingPrice As Variant
    OverrideReason As String
    ReferenceDistanceKm As Variant
    legacyMode As Boolean
    Option1Mode As Boolean
End Type


' ============================================================
' CALCULATION RESULT
' ============================================================

Public Type LineResult
    od As Variant
    ID As Variant
    MassPerM As Variant
    Metres As Variant
    MT As Variant
    Pipes As Variant
    Trailers As Variant
    FreightPerMT As Variant
    FCperMT As Variant
    ExMillPerMT As Variant
    BarePerM As Variant
    ExtPerM As Variant
    IntPerM As Variant
    FreightPerM As Variant
    TotalPerM As Variant
    TotalPerMT As Variant
    CalculatedPrice As Variant
    AppliedPrice As Variant
    PriceSource As String
    PriceVariance As Variant
    OverrideReason As String
    RawMaterialCostPerMT As Variant
    BarePipeCostPerMT As Variant
    LineValue As Variant
End Type


' ============================================================
' MAIN LINE CALCULATION
' ============================================================

Public Function CalcLine( _
    ByRef lineData As QuoteLineInput, _
    ByRef rates As RateInput, _
    ByVal priceByMT As Boolean _
) As LineResult

    On Error GoTo ErrorHandler

    Dim result As LineResult

    Dim od As Variant
    Dim idd As Variant
    Dim massKgPerM As Variant
    Dim massMtPerM As Variant

    Dim pre As Variant
    Dim ext As Variant
    Dim intc As Variant
    Dim freight As Variant
    Dim baseCost As Variant
    Dim financial As Variant

    Dim k As Variant

    ' --------------------------------------------------------
    ' Validate basic inputs
    ' --------------------------------------------------------

    If Len(Trim$(lineData.Basis)) = 0 Then
        Err.Raise vbObjectError + 1701, _
                  "CalcLine", _
                  "Basis is required."
    End If

    If Not IsNumeric(lineData.SizeMM) Then
        Err.Raise vbObjectError + 1702, _
                  "CalcLine", _
                  "Size must be numeric."
    End If

    If Not IsNumeric(lineData.WT) Then
        Err.Raise vbObjectError + 1703, _
                  "CalcLine", _
                  "Wall thickness must be numeric."
    End If

    If Not IsNumeric(lineData.Qty) Then
        Err.Raise vbObjectError + 1704, _
                  "CalcLine", _
                  "Quantity must be numeric."
    End If

    If Not IsNumeric(lineData.PipeLen) Then
        Err.Raise vbObjectError + 1705, _
                  "CalcLine", _
                  "Pipe length must be numeric."
    End If

    If Not IsNumeric(lineData.PerTrailer) Then
        Err.Raise vbObjectError + 1706, _
                  "CalcLine", _
                  "Pipes per trailer must be numeric."
    End If

    If CDec(lineData.WT) <= 0 Then
        Err.Raise vbObjectError + 1707, _
                  "CalcLine", _
                  "Wall thickness must be greater than zero."
    End If

    If CDec(lineData.Qty) < 0 Then
        Err.Raise vbObjectError + 1708, _
                  "CalcLine", _
                  "Quantity cannot be negative."
    End If

    If CDec(lineData.PipeLen) <= 0 Then
        Err.Raise vbObjectError + 1709, _
                  "CalcLine", _
                  "Pipe length must be greater than zero."
    End If

    If CDec(lineData.PerTrailer) <= 0 Then
        Err.Raise vbObjectError + 1715, _
                  "CalcLine", _
                  "Pipes per trailer must be greater than zero."
    End If


    ' --------------------------------------------------------
    ' GEOMETRY
    ' --------------------------------------------------------

    If UCase$(Trim$(lineData.Basis)) = "OD" Then

        od = CDec(lineData.SizeMM)

        idd = CDec( _
            CalculateID( _
                CDbl(od), _
                CDbl(lineData.WT), _
                "OD" _
            ) _
        )

    ElseIf UCase$(Trim$(lineData.Basis)) = "ID" Then

        idd = CDec(lineData.SizeMM)

        od = CDec( _
            CalculateOD( _
                CDbl(idd), _
                CDbl(lineData.WT), _
                "ID" _
            ) _
        )

    Else

        Err.Raise vbObjectError + 1710, _
                  "CalcLine", _
                  "Basis must be OD or ID."

    End If


    ' --------------------------------------------------------
    ' GEOMETRY VALIDATION
    ' --------------------------------------------------------

    If Not IsValidGeometry( _
        CDbl(od), _
        CDbl(idd), _
        CDbl(lineData.WT) _
    ) Then

        Err.Raise vbObjectError + 1711, _
                  "CalcLine", _
                  "Invalid pipe geometry."

    End If

    If CDec(idd) <= CDec(2) * CDec(lineData.WT) Then
        Err.Raise vbObjectError + 1718, _
                  "CalcLine", _
                  "Inner diameter must be greater than twice wall thickness."
    End If


    ' --------------------------------------------------------
    ' MASS
    ' --------------------------------------------------------

    massKgPerM = CDec( _
        CalculateMassKgPerM( _
            CDbl(od), _
            CDbl(lineData.WT) _
        ) _
    )
    massMtPerM = massKgPerM / CDec(1000)

    If massKgPerM <= 0 Then

        Err.Raise vbObjectError + 1716, _
                  "CalcLine", _
                  "Calculated mass per metre must be greater than zero."

    End If


    ' --------------------------------------------------------
    ' QUANTITY CONVERSION
    ' --------------------------------------------------------

    Select Case UCase$(Trim$(lineData.Unit))

        Case "METRES"

            result.Metres = CDec(lineData.Qty)

        Case "MT"

            result.Metres = _
                CDec(lineData.Qty) / _
                (massKgPerM / CDec(1000))

        Case "NOS"

            result.Metres = _
                CDec(lineData.Qty) * _
                CDec(lineData.PipeLen)

        Case Else

            Err.Raise vbObjectError + 1712, _
                      "CalcLine", _
                      "Unit must be Metres, MT, or Nos."

    End Select


    ' --------------------------------------------------------
    ' QUANTITY RESULTS
    ' --------------------------------------------------------

    result.MT = _
        result.Metres * _
        massKgPerM / _
        CDec(1000)

    result.Pipes = _
        RoundUpD( _
            result.Metres / CDec(lineData.PipeLen), _
            0 _
        )

    result.Trailers = _
        RoundUpD( _
            result.Pipes / CDec(lineData.PerTrailer), _
            0 _
        )


    ' --------------------------------------------------------
    ' PRE-COST
    ' --------------------------------------------------------

    pre = _
        CDec(rates.Steel) + _
        CDec(rates.GradeExtra)

    pre = _
        pre + _
        pre * CDec(rates.YieldPct) / CDec(100)

    pre = _
        pre + _
        CDec(rates.Conversion) + _
        CDec(rates.Testing) + _
        CDec(rates.Margin)


    ' --------------------------------------------------------
    ' EXTERNAL COATING
    ' --------------------------------------------------------

    If lineData.HasExt Then

        If rates.legacyMode Then
            ext = _
                CDec(rates.ExtRate) * _
                CDec(rates.PI) * _
                od / _
                CDec(1000) / _
                massMtPerM
        Else
            ext = _
                CDec(rates.ExtRate) * _
                CDec(rates.PI) * _
                od / _
                CDec(1000) / _
                massMtPerM
        End If

    Else

        ext = CDec(0)

    End If


    ' --------------------------------------------------------
    ' INTERNAL COATING
    ' --------------------------------------------------------

    If lineData.HasInt Then

        If rates.Option1Mode Then
            intc = _
                CDec(rates.IntRate) * _
                CDec(rates.PI) * _
                idd / _
                CDec(1000) / _
                massMtPerM
        Else
            intc = _
                CDec(rates.IntRate) * _
                CDec(rates.PI) * _
                od / _
                CDec(1000) / _
                massMtPerM
        End If

    Else

        intc = CDec(0)

    End If


    ' --------------------------------------------------------
    ' FREIGHT
    ' --------------------------------------------------------

    If result.MT <= 0 Then

        Err.Raise vbObjectError + 1717, _
                  "CalcLine", _
                  "Calculated MT must be greater than zero before freight calculation."

    End If

    If rates.Option1Mode Then
        freight = _
            RoundUpD( _
                CDec(rates.TripRate) * _
                CDec(result.Trailers) / _
                result.MT, _
                0 _
            ) + _
            CDec(rates.Handling)
    Else
        freight = _
            RoundUpD( _
                CDec(rates.TripRate) * _
                (CDec(result.Metres) / CDec(lineData.PipeLen)) / _
                CDec(lineData.PerTrailer) / _
                result.MT, _
                0 _
            ) + _
            CDec(rates.Handling)
    End If


    ' --------------------------------------------------------
    ' FINANCIAL COST
    ' --------------------------------------------------------

    If rates.Option1Mode Then

        baseCost = _
            pre + _
            ext + _
            intc + _
            freight

        k = _
            (CDec(rates.FinPct) / CDec(100)) * _
            (CDec(1) + CDec(rates.GSTPct) / CDec(100))

        If k <= 0 Or k >= 1 Then

            Err.Raise vbObjectError + 1713, _
                      "CalcLine", _
                      "Invalid financial/GST factor."

        End If

        financial = _
            k * _
            baseCost / _
            (CDec(1) - k)

    Else

        If CDec(rates.ManualFC) > 100 Then
            ' Explicit user manual override in Rs/MT (e.g. 1024, 1014, 1011, 1008 or legacy 1542, 1500)
            financial = CDec(rates.ManualFC)
        Else
            ' Approved formulaic financial cost:
            ' ROUND((PreCost + Ext + Int + Freight) * (1 + GST) * FinPct/100, 0)
            baseCost = pre + ext + intc + freight
            financial = Round( _
                baseCost * _
                (CDec(1) + CDec(rates.GSTPct) / CDec(100)) * _
                (CDec(rates.FinPct) / CDec(100)), _
                0 _
            )
        End If

    End If


    ' --------------------------------------------------------
    ' RESULT ASSIGNMENT
    ' --------------------------------------------------------

    result.od = od
    result.ID = idd

    result.MassPerM = massKgPerM

    result.FreightPerMT = freight

    result.FCperMT = financial

    result.ExMillPerMT = _
        pre + _
        financial + _
        ext + _
        intc


    ' --------------------------------------------------------
    ' SELLING PRICE PER METRE
    ' --------------------------------------------------------

    result.BarePerM = _
        RoundUpD( _
            (pre + financial) * _
            massMtPerM, _
            0 _
        )

    result.ExtPerM = _
        RoundUpD( _
            ext * massMtPerM, _
            0 _
        )

    result.IntPerM = _
        RoundUpD( _
            intc * massMtPerM, _
            0 _
        )

    result.FreightPerM = _
        RoundUpD( _
            freight * massMtPerM, _
            0 _
        )

    result.TotalPerM = _
        result.BarePerM + _
        result.ExtPerM + _
        result.IntPerM + _
        result.FreightPerM


    ' --------------------------------------------------------
    ' SELLING PRICE PER MT
    ' --------------------------------------------------------

    result.TotalPerMT = _
        RoundUpD(pre + financial, 0) + _
        RoundUpD(ext, 0) + _
        RoundUpD(intc, 0) + _
        result.FreightPerMT


    ' --------------------------------------------------------
    ' COMMERCIAL LINE VALUE & APPLIED SELLING PRICE
    ' --------------------------------------------------------

    Dim calcPrice As Variant
    Dim appliedPrice As Variant
    Dim priceSrc As String
    Dim priceVar As Variant

    If priceByMT Then
        calcPrice = result.TotalPerMT
    Else
        calcPrice = result.TotalPerM
    End If

    If CDec(rates.ManualSellingPrice) > 0 Then
        appliedPrice = CDec(rates.ManualSellingPrice)
        priceSrc = "MANUAL_OVERRIDE"
        priceVar = appliedPrice - calcPrice
    Else
        appliedPrice = calcPrice
        priceSrc = "CALCULATED"
        priceVar = CDec(0)
    End If

    result.CalculatedPrice = calcPrice
    result.AppliedPrice = appliedPrice
    result.PriceSource = priceSrc
    result.PriceVariance = priceVar
    result.OverrideReason = rates.OverrideReason
    result.RawMaterialCostPerMT = (CDec(rates.Steel) + CDec(rates.GradeExtra)) * (CDec(1) + CDec(rates.YieldPct) / CDec(100))
    result.BarePipeCostPerMT = pre

    If priceByMT Then
        result.LineValue = appliedPrice * result.MT
    Else
        result.LineValue = appliedPrice * result.Metres
    End If

    CalcLine = result

    Exit Function


ErrorHandler:

    Err.Raise _
        Err.Number, _
        "CalcLine", _
        Err.Description

End Function


' ============================================================
' SETTINGS LOOKUP
' ============================================================

Public Function SettingValue(ByVal key As String) As Variant

    On Error GoTo ErrorHandler

    Dim settings As ListObject
    Dim rowObject As ListRow

    Set settings = _
        ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")

    For Each rowObject In settings.ListRows

        If StrComp( _
            CStr( _
                rowObject.Range.Cells( _
                    1, _
                    settings.ListColumns("SettingKey").index _
                ).Value2 _
            ), _
            key, _
            vbTextCompare _
        ) = 0 Then

            SettingValue = _
                rowObject.Range.Cells( _
                    1, _
                    settings.ListColumns("SettingValue").index _
                ).Value2

            Exit Function

        End If

    Next rowObject

    Err.Raise _
        vbObjectError + 1714, _
        "SettingValue", _
        "Setting not found: " & key

    Exit Function


ErrorHandler:

    Err.Raise _
        Err.Number, _
        "SettingValue", _
        Err.Description

End Function


' ============================================================
' LOAD DEFAULT RATE INPUT
' ============================================================

Public Sub LoadDefaultRateInput( _
    ByRef rates As RateInput, _
    ByVal legacyMode As Boolean _
)

    On Error GoTo ErrorHandler

    With rates

        .Steel = _
            CDec(SettingValue("DefaultSteelRate"))

        .GradeExtra = _
            CDec(SettingValue("DefaultGradeExtra"))

        .YieldPct = _
            CDec(SettingValue("DefaultYieldPct"))

        .Conversion = _
            CDec(SettingValue("DefaultConversion"))

        .Testing = _
            CDec(SettingValue("DefaultTesting"))

        .Margin = _
            CDec(SettingValue("DefaultMargin"))

        .FinPct = _
            CDec(SettingValue("DefaultFinancialPct"))

        .GSTPct = _
            CDec(SettingValue("GSTPct"))

        .PI = _
            CDec(SettingValue("PIPct"))

        .Handling = _
            CDec(SettingValue("HandlingRsPerMT"))


        .legacyMode = legacyMode
        .Option1Mode = False

        .TripRate = _
            CDec(SettingValue("TripRateManual"))

    End With

    Exit Sub


ErrorHandler:

    Err.Raise _
        Err.Number, _
        "LoadDefaultRateInput", _
        Err.Description

End Sub


' ============================================================
' P3 GOLDEN TOTAL
' ============================================================

Public Function P3GoldenTotal(ByVal legacy As Boolean) As Variant

    On Error GoTo ErrorHandler

    Dim result As LineResult
    Dim index As Long
    Dim total As Variant
    total = CDec(0)
    For index = 0 To 3
        result = P3GoldenLine(legacy, index)
        total = total + result.LineValue
    Next index
    P3GoldenTotal = total
    Exit Function
ErrorHandler:
    Err.Raise _
        Err.Number, _
        "P3GoldenTotal", _
        Err.Description
End Function

Public Function P3GoldenLine(ByVal legacy As Boolean, ByVal index As Long) As LineResult
    Dim ids As Variant
    Dim lengths As Variant
    Dim manualFinancial As Variant
    Dim rates As RateInput
    Dim lineData As QuoteLineInput

    ids = Array(1226, 1300, 1350, 1400)
    lengths = Array(2000, 350, 500, 4400)
    manualFinancial = Array(1542, 1534, 1529, 1525)
    If index < 0 Or index > 3 Then
        Err.Raise vbObjectError + 1804, "P3GoldenLine", "Golden line index out of range."
    End If

    LoadDefaultRateInput rates, legacy
    rates.Option1Mode = Not legacy
    rates.Steel = CDec(61000)
    rates.Conversion = CDec(5000)
    rates.FinPct = CDec(1.5)
    rates.ExtRate = CDec(500)
    rates.IntRate = CDec(275)
    If legacy Then
        rates.TripRate = CDec(98000)
        ' The reference workbook uses a 3.14 circumference constant.
        rates.PI = CDec(3.14)
    Else
        rates.TripRate = CDec(125.32) * CDec(782)
    End If

    With lineData
        .Basis = "ID"
        .SizeMM = CDec(ids(index))
        .WT = CDec(8.4)
        .Qty = CDec(lengths(index))
        .Unit = "Metres"
        .PipeLen = CDec(12)
        .PerTrailer = CDec(4)
        .HasExt = True
        .HasInt = True
    End With
    rates.ManualFC = CDec(manualFinancial(index))
    P3GoldenLine = CalcLine(lineData, rates, False)
End Function


' ============================================================
' APPROVED COMPANY COSTSHEET GOLDEN REFERENCE
' ============================================================

Public Function ApprovedGoldenTotal() As Variant
    On Error GoTo ErrorHandler
    Dim result As LineResult
    Dim index As Long
    Dim total As Variant
    total = CDec(0)
    For index = 0 To 3
        result = ApprovedGoldenLine(index)
        total = total + result.LineValue
    Next index
    ApprovedGoldenTotal = total
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "ApprovedGoldenTotal", Err.Description
End Function

Public Function ApprovedGoldenLine(ByVal index As Long) As LineResult
    Dim ids As Variant
    Dim lengths As Variant
    Dim manualFinancial As Variant
    Dim rates As RateInput
    Dim lineData As QuoteLineInput

    ids = Array(1226, 1300, 1350, 1400)
    lengths = Array(2000, 350, 500, 4400)
    manualFinancial = Array(1024, 1014, 1011, 1008)
    If index < 0 Or index > 3 Then
        Err.Raise vbObjectError + 1805, "ApprovedGoldenLine", "Approved line index out of range."
    End If

    LoadDefaultRateInput rates, False
    rates.Option1Mode = False
    rates.ExtRate = CDec(500)
    rates.IntRate = CDec(250)
    rates.TripRate = CDec(SettingValue("TripRateManual"))
    rates.PI = CDec(3.14)

    With lineData
        .Basis = "ID"
        .SizeMM = CDec(ids(index))
        .WT = CDec(8.4)
        .Qty = CDec(lengths(index))
        .Unit = "Metres"
        .PipeLen = CDec(12)
        .PerTrailer = CDec(4)
        .HasExt = True
        .HasInt = True
    End With
    rates.ManualFC = CDec(manualFinancial(index))
    ApprovedGoldenLine = CalcLine(lineData, rates, False)
End Function
