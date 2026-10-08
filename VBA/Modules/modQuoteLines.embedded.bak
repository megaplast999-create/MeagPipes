Attribute VB_Name = "modQuoteLines"
Option Explicit

Public Sub AddQuoteLine(ByVal quoteId As String, ByVal productId As String, ByVal family As String, _
    ByVal basis As String, ByVal sizeMm As Double, ByVal thicknessMm As Double, _
    ByVal gradeId As String, ByVal quantity As Double, ByVal quantityUnit As String, _
    ByVal pipeLengthM As Double, ByVal plantId As String)

    Dim tableObject As ListObject
    Dim lineObject As ListRow
    Set tableObject = ThisWorkbook.Worksheets("QUOTE_LINES").ListObjects("tblQuoteLines")

    If Len(Trim$(quoteId)) = 0 Or Len(Trim$(productId)) = 0 Or Len(Trim$(family)) = 0 Or Len(Trim$(gradeId)) = 0 Then
        Err.Raise vbObjectError + 1301, "AddQuoteLine", "Quote, product, family, and grade are required."
    End If
    If Not IsValidQuantity(quantity) Then
        Err.Raise vbObjectError + 1302, "AddQuoteLine", "Quantity must be greater than zero."
    End If
    If Not IsValidPipeLength(pipeLengthM) Then
        Err.Raise vbObjectError + 1303, "AddQuoteLine", "Pipe length must be between 6 and 13 metres."
    End If
    If UCase$(Trim$(basis)) <> "OD" And UCase$(Trim$(basis)) <> "ID" Then
        Err.Raise vbObjectError + 1304, "AddQuoteLine", "Basis must be OD or ID."
    End If

    Set lineObject = tableObject.ListRows.Add
    With lineObject.Range
        .Cells(1, tableObject.ListColumns("LineID").Index).value = "L" & Format$(tableObject.ListRows.Count, "0000")
        .Cells(1, tableObject.ListColumns("QuoteID").Index).value = quoteId
        .Cells(1, tableObject.ListColumns("ProductID").Index).value = productId
        .Cells(1, tableObject.ListColumns("Family").Index).value = family
        .Cells(1, tableObject.ListColumns("Basis").Index).value = basis
        .Cells(1, tableObject.ListColumns("SizeMM").Index).value = sizeMm
        .Cells(1, tableObject.ListColumns("ThicknessMM").Index).value = thicknessMm
        .Cells(1, tableObject.ListColumns("GradeID").Index).value = gradeId
        .Cells(1, tableObject.ListColumns("Quantity").Index).value = quantity
        .Cells(1, tableObject.ListColumns("QuantityUnit").Index).value = quantityUnit
        .Cells(1, tableObject.ListColumns("PipeLengthM").Index).value = pipeLengthM
        .Cells(1, tableObject.ListColumns("PlantID").Index).value = plantId
        .Cells(1, tableObject.ListColumns("Status").Index).value = "DRAFT"
    End With
End Sub
