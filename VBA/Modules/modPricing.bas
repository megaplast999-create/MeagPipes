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
    ' Cost policy (set from SETTINGS by ApplyCostPolicy). All three default to
    ' False = the reference cost-sheet behaviour of the old non-Option1 path,
    ' so code that never calls ApplyCostPolicy is unchanged.
    WholeTrailers As Boolean        ' True: CEILING trailers; False: fractional trips (reference sheet)
    IntCoatOnID As Boolean          ' True: inside surface (ID); False: OD factor (reference sheet)
    FCAlgebraic As Boolean          ' True: FC = k*base/(1-k); False: old ROUND(base*(1+GST)*fin)
    TripsOverride As Variant        ' optional manual loading-plan trips for THIS line (Empty = off)
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
    FreightTrips As Variant         ' trips the freight Rs/MT was actually built on
    ExtPerMT As Variant             ' external coating Rs/MT (unrounded)
    IntPerMT As Variant             ' internal coating Rs/MT (unrounded)
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

        If rates.Option1Mode Or rates.IntCoatOnID Then
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

    ' Trips the freight is built on:
    '   manual loading-plan trips (TripsOverride)        - if the user keyed them
    '   else whole trailers  = CEILING(pipes / per trailer)
    '   else fractional trips = (metres / pipe length) / per trailer   [reference sheet]
    Dim costedTrips As Variant
    costedTrips = Empty
    If Not IsEmpty(rates.TripsOverride) Then
        If IsNumeric(rates.TripsOverride) Then
            If CDec(rates.TripsOverride) > 0 Then costedTrips = CDec(rates.TripsOverride)
        End If
    End If
    If IsEmpty(costedTrips) Then
        If rates.Option1Mode Or rates.WholeTrailers Then
            costedTrips = CDec(result.Trailers)
        Else
            costedTrips = (CDec(result.Metres) / CDec(lineData.PipeLen)) / _
                          CDec(lineData.PerTrailer)
        End If
    End If

    freight = _
        RoundUpD( _
            CDec(rates.TripRate) * _
            costedTrips / _
            result.MT, _
            0 _
        ) + _
        CDec(rates.Handling)
    result.FreightTrips = costedTrips


    ' --------------------------------------------------------
    ' FINANCIAL COST
    ' --------------------------------------------------------

    If rates.Option1Mode Or rates.FCAlgebraic Then

        ' A typed FC wins under the explicit policy (cost-sheet row 32).
        ' The old Option1 contract keeps ignoring it.
        If rates.FCAlgebraic And Not rates.Option1Mode And CDec(rates.ManualFC) > 0 Then

            financial = CDec(rates.ManualFC)

        Else

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

            ' Fixed point of the cost sheet's own check cell (row 41):
            ' FC = fin% * (1+GST) * (base + FC)  =>  FC = k*base/(1-k)
            financial = _
                k * _
                baseCost / _
                (CDec(1) - k)

        End If

    Else

        If CDec(rates.ManualFC) > 100 Then
            ' Explicit user manual override in Rs/MT (e.g. 1024, 1014, 1011, 1008 or legacy 1542, 1500)
            financial = CDec(rates.ManualFC)
        Else
            ' Old simple formula (leaves FC itself out of the base):
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
    result.ExtPerMT = ext
    result.IntPerMT = intc

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


' ============================================================
' COST POLICY + SINGLE LINE-INPUT LOADER
'
' Every place that prices a quote line (AddQuoteLine,
' UpdateQuoteLineFull, CalcAllLines, WriteCostingSheet) used to carry its
' own 60-line copy of "build the rates", each with its own hard-coded
' 500 / 250 / 3.14 / 4-pipes-per-trailer and different override handling.
' They disagreed (e.g. editing a line dropped its typed FC), so the same
' line priced differently depending on the button pressed.
' LoadLineInputsFromRow is now the only place that does this.
' ============================================================

Public Function SettingValueOr(ByVal key As String, ByVal defaultValue As Variant) As Variant
    Dim v As Variant
    On Error Resume Next
    v = SettingValue(key)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        SettingValueOr = defaultValue
        Exit Function
    End If
    On Error GoTo 0
    If IsNull(v) Then
        SettingValueOr = defaultValue
    ElseIf IsEmpty(v) Then
        SettingValueOr = defaultValue
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        SettingValueOr = defaultValue
    Else
        SettingValueOr = v
    End If
End Function

' Reads the three cost-policy switches from SETTINGS (defaults = reference cost sheet).
Public Sub ApplyCostPolicy(ByRef rates As RateInput)
    rates.WholeTrailers = (UCase$(Trim$(CStr(SettingValueOr("FreightTrailerRounding", "FRACTIONAL")))) = "WHOLE")
    rates.IntCoatOnID = (UCase$(Trim$(CStr(SettingValueOr("IntCoatingBasis", "OD")))) = "ID")
    rates.FCAlgebraic = (UCase$(Trim$(CStr(SettingValueOr("FCMethod", "ALGEBRAIC")))) = "ALGEBRAIC")
End Sub

' Adds the new SETTINGS rows if the workbook does not have them yet. Idempotent.
Public Sub EnsureCostPolicySettings()
    Dim tbl As ListObject
    Dim keys As Variant, vals As Variant, descs As Variant
    Dim i As Long
    Dim rowObj As ListRow
    Dim found As Boolean
    Dim newRow As ListRow

    On Error GoTo Skip
    Set tbl = ThisWorkbook.Worksheets("SETTINGS").ListObjects("tblSettings")
    keys = Array("FreightTrailerRounding", "IntCoatingBasis", "FCMethod", "DefaultExtCoatingRate", "DefaultIntCoatingRate")
    vals = Array("FRACTIONAL", "OD", "ALGEBRAIC", 500, 250)
    descs = Array("Freight trips: FRACTIONAL (reference cost sheet) or WHOLE (round trailers up)", _
                  "Internal coating area basis: OD (reference cost sheet) or ID (true inside surface)", _
                  "Financial cost: ALGEBRAIC (self-consistent, matches sheet check cell) or SIMPLE (old shortcut)", _
                  "Fallback external coating rate Rs/m2 when the coating master has no rate", _
                  "Fallback internal coating rate Rs/m2 when the coating master has no rate")
    For i = 0 To UBound(keys)
        found = False
        For Each rowObj In tbl.ListRows
            If StrComp(CStr(rowObj.Range.Cells(1, tbl.ListColumns("SettingKey").Index).Value), CStr(keys(i)), vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next rowObj
        If Not found Then
            Set newRow = tbl.ListRows.Add
            newRow.Range.Cells(1, tbl.ListColumns("SettingKey").Index).Value = keys(i)
            newRow.Range.Cells(1, tbl.ListColumns("SettingValue").Index).Value = vals(i)
            newRow.Range.Cells(1, tbl.ListColumns("Description").Index).Value = descs(i)
            newRow.Range.Cells(1, tbl.ListColumns("DefaultValue").Index).Value = vals(i)
            newRow.Range.Cells(1, tbl.ListColumns("Editable").Index).Value = "Yes"
            newRow.Range.Cells(1, tbl.ListColumns("Category").Index).Value = "Pricing"
        End If
    Next i
Skip:
End Sub

Private Function ColIdx(ByVal tbl As ListObject, ByVal colName As String) As Long
    Dim c As ListColumn
    ColIdx = 0
    For Each c In tbl.ListColumns
        If StrComp(c.Name, colName, vbTextCompare) = 0 Then
            ColIdx = c.Index
            Exit Function
        End If
    Next c
End Function

Private Function RowNum(ByVal tbl As ListObject, ByVal lineRow As ListRow, ByVal colName As String) As Double
    Dim ci As Long
    ci = ColIdx(tbl, colName)
    If ci > 0 Then RowNum = Val(lineRow.Range.Cells(1, ci).Value) Else RowNum = 0#
End Function

Private Function RowText(ByVal tbl As ListObject, ByVal lineRow As ListRow, ByVal colName As String) As String
    Dim ci As Long
    ci = ColIdx(tbl, colName)
    If ci > 0 Then RowText = Trim$(CStr(lineRow.Range.Cells(1, ci).Value)) Else RowText = ""
End Function

' Rs/m2 for a selected coating from the MASTERS coating table; fallback when blank/zero.
Public Function CoatingRateFor(ByVal tableName As String, ByVal coatingId As String, ByVal fallbackRate As Double) As Double
    Dim tbl As ListObject
    Dim rowObj As ListRow
    Dim rate As Double

    CoatingRateFor = fallbackRate
    If Len(Trim$(coatingId)) = 0 Then Exit Function
    On Error GoTo Done
    Set tbl = ThisWorkbook.Worksheets("MASTERS").ListObjects(tableName)
    For Each rowObj In tbl.ListRows
        If StrComp(Trim$(CStr(rowObj.Range.Cells(1, tbl.ListColumns("CoatingID").Index).Value)), coatingId, vbTextCompare) = 0 Then
            rate = Val(rowObj.Range.Cells(1, tbl.ListColumns("RateRsPerM2").Index).Value)
            If rate > 0 Then CoatingRateFor = rate
            Exit Function
        End If
    Next rowObj
Done:
End Function

' Which coatings does the line carry? The stored coating selection decides;
' only when both are blank does the product (P02/P04 = coated) decide.
Public Sub LineCoatingFlags(ByVal productId As String, ByVal extId As String, ByVal intId As String, _
                            ByRef hasExt As Boolean, ByRef hasInt As Boolean)
    Dim extSel As Boolean, intSel As Boolean
    extSel = (Len(extId) > 0 And UCase$(extId) <> "BARE" And UCase$(extId) <> "NONE")
    intSel = (Len(intId) > 0 And UCase$(intId) <> "BARE" And UCase$(intId) <> "NONE")
    If extSel Or intSel Then
        hasExt = extSel
        hasInt = intSel
    Else
        hasExt = (InStr(1, productId, "02", vbTextCompare) > 0 Or InStr(1, productId, "04", vbTextCompare) > 0)
        hasInt = hasExt
    End If
End Sub

' Builds QuoteLineInput + RateInput for one QUOTE_LINES row. Callers may overwrite
' individual rates afterwards (e.g. a manual selling price typed in the form).
Public Sub LoadLineInputsFromRow(ByRef lineData As QuoteLineInput, ByRef rates As RateInput, _
                                 ByVal tbl As ListObject, ByVal lineRow As ListRow, _
                                 ByVal legacyMode As Boolean)
    Dim basisVal As String
    Dim sizeVal As Double, wtVal As Double, pipeLen As Double
    Dim odMm As Double, massKg As Double
    Dim prodId As String, extId As String, intId As String
    Dim hasExt As Boolean, hasInt As Boolean
    Dim perTrailer As Long
    Dim v As Double

    basisVal = RowText(tbl, lineRow, "Basis")
    sizeVal = RowNum(tbl, lineRow, "SizeMM")
    wtVal = RowNum(tbl, lineRow, "ThicknessMM")
    pipeLen = RowNum(tbl, lineRow, "PipeLengthM")
    If pipeLen < 6# Or pipeLen > 13# Then pipeLen = 12#
    prodId = RowText(tbl, lineRow, "ProductID")
    extId = RowText(tbl, lineRow, "CoatingExternalID")
    intId = RowText(tbl, lineRow, "CoatingInternalID")

    LineCoatingFlags prodId, extId, intId, hasExt, hasInt

    ' Pipes per trailer from the stacking master (tblStacking), else payload / pipe weight.
    perTrailer = 4
    On Error Resume Next
    odMm = CalculateOD(sizeVal, wtVal, basisVal)
    massKg = CalculateMassKgPerM(odMm, wtVal)
    If Err.Number = 0 And odMm > 0 And massKg > 0 Then
        perTrailer = GetPipesPerTrailer(odMm, massKg * pipeLen / 1000#)
    End If
    Err.Clear
    On Error GoTo 0
    If perTrailer <= 0 Then perTrailer = 4

    With lineData
        .Basis = basisVal
        .SizeMM = CDec(sizeVal)
        .WT = CDec(wtVal)
        .Qty = CDec(RowNum(tbl, lineRow, "Quantity"))
        .Unit = RowText(tbl, lineRow, "QuantityUnit")
        .PipeLen = CDec(pipeLen)
        .PerTrailer = CDec(perTrailer)
        .HasExt = hasExt
        .HasInt = hasInt
    End With

    LoadDefaultRateInput rates, legacyMode
    rates.Option1Mode = False
    rates.TripsOverride = Empty      ' rates is reused line after line - never inherit the previous line's trips
    ApplyCostPolicy rates
    rates.PI = CDec(SettingValueOr("PIPct", 3.14))
    rates.ExtRate = CDec(CoatingRateFor("tblExternalCoatings", extId, CDbl(SettingValueOr("DefaultExtCoatingRate", 500))))
    rates.IntRate = CDec(CoatingRateFor("tblInternalCoatings", intId, CDbl(SettingValueOr("DefaultIntCoatingRate", 250))))
    rates.TripRate = CDec(SettingValue("TripRateManual"))
    rates.ManualFC = CDec(0)
    rates.ManualSellingPrice = CDec(0)
    rates.OverrideReason = ""

    ' Optional per-line typed values (blank = use master / settings)
    v = RowNum(tbl, lineRow, "ManualFC")
    If v <= 0 Then v = RowNum(tbl, lineRow, "FCOverride")
    If v > 0 Then rates.ManualFC = CDec(v)
    v = RowNum(tbl, lineRow, "SteelRate"):          If v > 0 Then rates.Steel = CDec(v)
    v = RowNum(tbl, lineRow, "YieldPct"):           If v > 0 Then rates.YieldPct = CDec(v)
    v = RowNum(tbl, lineRow, "ConversionRate"):     If v > 0 Then rates.Conversion = CDec(v)
    v = RowNum(tbl, lineRow, "ExtCoatingRate"):     If v > 0 Then rates.ExtRate = CDec(v)
    v = RowNum(tbl, lineRow, "IntCoatingRate"):     If v > 0 Then rates.IntRate = CDec(v)
    v = RowNum(tbl, lineRow, "TripRate"):           If v > 0 Then rates.TripRate = CDec(v)
    v = RowNum(tbl, lineRow, "FinPct"):             If v > 0 Then rates.FinPct = CDec(v)
    v = RowNum(tbl, lineRow, "ManualSellingPrice"): If v > 0 Then rates.ManualSellingPrice = CDec(v)
    v = RowNum(tbl, lineRow, "ManualTrips"):        If v > 0 Then rates.TripsOverride = CDec(v)
    rates.OverrideReason = RowText(tbl, lineRow, "OverrideReason")
End Sub
