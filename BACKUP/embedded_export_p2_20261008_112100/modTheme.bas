Attribute VB_Name = "modTheme"
Option Explicit

Private Const NAVY As String = "1F4E78"
Private Const LIGHT_BLUE As String = "D9EAF7"
Private Const LIGHT_GREY As String = "F3F6F9"
Private Const COMPETITOR_YELLOW As String = "FFF2CC"

Public Function HexToLong(ByVal hexColor As String) As Long
    Dim value As String
    Dim red As Long
    Dim green As Long
    Dim blue As Long

    value = Replace$(Trim$(hexColor), "#", vbNullString)
    If Len(value) <> 6 Or Not IsHex(value) Then
        Err.Raise vbObjectError + 1601, "HexToLong", _
            "Colour must be a six-digit hexadecimal value."
    End If

    red = CLng("&H" & Mid$(value, 1, 2))
    green = CLng("&H" & Mid$(value, 3, 2))
    blue = CLng("&H" & Mid$(value, 5, 2))
    HexToLong = RGB(red, green, blue)
End Function

Public Sub ColourRow(ByVal target As Range, ByVal fillHex As String, _
                     Optional ByVal fontHex As String = "FFFFFF")
    If target Is Nothing Then Exit Sub
    With target
        .Interior.Pattern = xlSolid
        .Interior.Color = HexToLong(fillHex)
        .Font.Color = HexToLong(fontHex)
    End With
End Sub

Public Sub ApplySheetTheme(ByVal ws As Worksheet)
    Dim used As Range
    Dim tableObject As ListObject
    Dim header As Range

    If ws Is Nothing Then Exit Sub
    Set used = ws.UsedRange
    If used Is Nothing Then Exit Sub

    With used
        .Font.Name = "Aptos"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With

    ColourRow used.Rows(1), NAVY

    For Each tableObject In ws.ListObjects
        Set header = tableObject.HeaderRowRange
        ColourRow header, NAVY
        header.Font.Bold = True
        header.HorizontalAlignment = xlCenter
        If Not tableObject.DataBodyRange Is Nothing Then
            tableObject.DataBodyRange.Borders.LineStyle = xlContinuous
            tableObject.DataBodyRange.Borders.Color = HexToLong("D9E2F3")
        End If
    Next tableObject

    Select Case UCase$(ws.Name)
        Case "START", "QUOTE", "COSTING", "OFFER"
            StyleSectionLabels ws
        Case "QUOTE_LINES"
            ws.Rows(3).RowHeight = 24
    End Select

    ApplyCompetitorHighlights ws
End Sub

Public Sub ApplyCompetitorHighlights(ByVal ws As Worksheet)
    Dim rowObject As Range
    Dim cell As Range
    Dim marker As String

    If ws Is Nothing Then Exit Sub
    For Each rowObject In ws.UsedRange.Rows
        marker = vbNullString
        For Each cell In rowObject.Cells
            If VarType(cell.Value2) = vbString Then
                If InStr(1, CStr(cell.Value2), "COMPETITOR", vbTextCompare) > 0 Then
                    marker = "COMPETITOR"
                    Exit For
                End If
            End If
        Next cell
        If Len(marker) > 0 Then ColourRow rowObject, COMPETITOR_YELLOW, "7F6000"
    Next rowObject
End Sub

Private Sub StyleSectionLabels(ByVal ws As Worksheet)
    Dim cell As Range
    For Each cell In ws.UsedRange.Cells
        If Len(Trim$(CStr(cell.Value2))) > 0 Then
            If UCase$(CStr(cell.Value2)) = UCase$(CStr(cell.Value2)) _
                And (InStr(1, CStr(cell.Value2), "QUOTE", vbTextCompare) > 0 _
                Or InStr(1, CStr(cell.Value2), "COST", vbTextCompare) > 0 _
                Or InStr(1, CStr(cell.Value2), "OFFER", vbTextCompare) > 0 _
                Or InStr(1, CStr(cell.Value2), "STATUS", vbTextCompare) > 0) Then
                cell.Font.Bold = True
            End If
        End If
    Next cell
End Sub

Private Function IsHex(ByVal value As String) As Boolean
    Dim index As Long
    Dim character As String
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If InStr(1, "0123456789ABCDEFabcdef", character, vbBinaryCompare) = 0 Then Exit Function
    Next index
    IsHex = True
End Function
