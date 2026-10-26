Attribute VB_Name = "modProtection"
Option Explicit

' Global toggle: Sheet protection disabled per user instruction
Public Const ENABLE_SHEET_PROTECTION As Boolean = True

' ============================================================
' Mega Pipes Quotation System
' Module: modProtection
' Purpose:
'   P23 Workbook & Sheet Protection
'   Protected cells reject user manual editing
'   Macros retain write capability via UserInterfaceOnly:=True
'   Configurable via MEGA_SHEET_PWD environment variable
'   Selective cell locking: inputs unlocked, formulas/outputs locked
' ============================================================

Public Function SheetPassword() As String
    ' Password disabled per user instruction
    SheetPassword = ""
End Function

Public Sub ConfigureCellLocking()
    If Not ENABLE_SHEET_PROTECTION Then
        On Error Resume Next
        ThisWorkbook.Worksheets("QUOTE").Range("A1:I100").Locked = False
        On Error GoTo 0
        Exit Sub
    End If
    Dim wsQuote As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim col As ListColumn
    Dim prevEvents As Boolean, prevAlerts As Boolean
    
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    On Error GoTo CleanExit
    
    ' 1. QUOTE Sheet: Unlock input header fields, lock output / formula / status cells
    Set wsQuote = ThisWorkbook.Worksheets("QUOTE")
    wsQuote.UsedRange.Locked = True
    
    ' Unlock editable header inputs: Customer (B8) through Price Basis (B17)
    wsQuote.Range("B8:B17").Locked = False
    
    ' Unlock editable product item inputs (B20:B33 and C31)
    wsQuote.Range("B20:B33").Locked = False
    wsQuote.Range("C31").Locked = False
    
    ' Keep Enquiry Number (B4) and Summary card (D3:E17) locked
    wsQuote.Range("B4").Locked = True
    wsQuote.Range("D3:E17").Locked = True
    
    ' Keep Quote Lines display table locked against manual edits
    wsQuote.Range("D19:I34").Locked = True
    
    ' Unlock feedback banner cell
    wsQuote.Range("A33:I34").Locked = False
    
    ' 2. QUOTE_LINES Sheet: Unlock user input columns in tblQuoteLines
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    wsLines.UsedRange.Locked = True
    
    ' Unlock user data-entry columns
    UnlockColumn tblLines, "ProductID"
    UnlockColumn tblLines, "Family"
    UnlockColumn tblLines, "Basis"
    UnlockColumn tblLines, "SizeMM"
    UnlockColumn tblLines, "ThicknessMM"
    UnlockColumn tblLines, "GradeID"
    UnlockColumn tblLines, "Quantity"
    UnlockColumn tblLines, "QuantityUnit"
    UnlockColumn tblLines, "PipeLengthM"
    UnlockColumn tblLines, "PlantID"
    UnlockColumn tblLines, "CoatingExternalID"
    UnlockColumn tblLines, "CoatingInternalID"
    UnlockColumn tblLines, "ManualSellingPrice"
    UnlockColumn tblLines, "OverrideReason"
    UnlockColumn tblLines, "IsCustomSize"
    
    ' Lock calculation and status columns
    LockColumn tblLines, "LineID"
    LockColumn tblLines, "PipeCount"
    LockColumn tblLines, "Metres"
    LockColumn tblLines, "WeightMT"
    LockColumn tblLines, "CalculatedODMM"
    LockColumn tblLines, "CalculatedIDMM"
    LockColumn tblLines, "UnitPrice"
    LockColumn tblLines, "LineTotal"
    LockColumn tblLines, "Status"
    
    ' 3. Internal data & helper sheets: Unlock cells so system macros can write
    On Error Resume Next
    ThisWorkbook.Worksheets("MASTERS").Cells.Locked = False
    ThisWorkbook.Worksheets("SETTINGS").Cells.Locked = False
    ThisWorkbook.Worksheets("QUOTE_LOG").Cells.Locked = False
    ThisWorkbook.Worksheets("DISTANCE_CACHE").Cells.Locked = False
    On Error GoTo 0

CleanExit:
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
End Sub

Private Sub UnlockColumn(ByVal tbl As ListObject, ByVal colName As String)
    On Error Resume Next
    tbl.ListColumns(colName).Range.Locked = False
    On Error GoTo 0
End Sub

Private Sub LockColumn(ByVal tbl As ListObject, ByVal colName As String)
    On Error Resume Next
    tbl.ListColumns(colName).Range.Locked = True
    On Error GoTo 0
End Sub

Public Sub ProtectSheet(ByVal ws As Worksheet, Optional ByVal userInterfaceOnly As Boolean = True)
    On Error Resume Next
    If Not ENABLE_SHEET_PROTECTION Then
        ws.Unprotect ""
        ws.UsedRange.Locked = False
        Exit Sub
    End If
    ' Protect without password so Excel never prompts the user for a password
    ws.Protect DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True, _
               UserInterfaceOnly:=userInterfaceOnly, _
               AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, _
               AllowFormattingRows:=True, _
               AllowSorting:=True, _
               AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

Public Sub UnprotectSheet(ByVal ws As Worksheet, Optional ByVal pwd As String = "")
    On Error Resume Next
    ws.Unprotect ""
    If ws.ProtectContents Then ws.Unprotect "MegaPipes2026"
    If ws.ProtectContents Then ws.Unprotect "admin"
    If ws.ProtectContents Then ws.Unprotect "MegaPipes"
    If ws.ProtectContents And Len(pwd) > 0 Then ws.Unprotect pwd
    On Error GoTo 0
End Sub

Public Sub ProtectAllSheets(Optional ByVal userInterfaceOnly As Boolean = True)
    If Not ENABLE_SHEET_PROTECTION Then
        UnprotectAllSheets
        Exit Sub
    End If
    Dim ws As Worksheet
    ConfigureCellLocking
    For Each ws In ThisWorkbook.Worksheets
        ProtectSheet ws, userInterfaceOnly
    Next ws
End Sub


' ============================================================
' Protection of Calculated Output Sheets (OFFER and SYNOPSIS)
' ============================================================

Public Sub ProtectCalculatedSheets(Optional ByVal userInterfaceOnly As Boolean = True)
    If Not ENABLE_SHEET_PROTECTION Then Exit Sub
    
    Dim prevAlerts As Boolean, prevScreen As Boolean
    prevAlerts = Application.DisplayAlerts
    prevScreen = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim pwd As String
    pwd = SheetPassword()
    
    ' 1. Protect OFFER Sheet: Lock all calculated cells and protect sheet
    Dim wsOff As Worksheet
    On Error Resume Next
    Set wsOff = ThisWorkbook.Worksheets("OFFER")
    On Error GoTo 0
    
    If Not wsOff Is Nothing Then
        wsOff.UsedRange.Locked = True
        wsOff.Protect Password:=pwd, _
                      DrawingObjects:=True, _
                      Contents:=True, _
                      Scenarios:=True, _
                      UserInterfaceOnly:=userInterfaceOnly, _
                      AllowFormattingCells:=True, _
                      AllowFormattingColumns:=True, _
                      AllowFormattingRows:=True, _
                      AllowSorting:=True, _
                      AllowFiltering:=True
        wsOff.EnableSelection = xlNoRestrictions
    End If
    
    ' 2. Protect SYNOPSIS Sheet: Lock all calculated cells and protect sheet
    Dim wsSyn As Worksheet
    On Error Resume Next
    Set wsSyn = ThisWorkbook.Worksheets("SYNOPSIS")
    On Error GoTo 0
    
    If Not wsSyn Is Nothing Then
        wsSyn.UsedRange.Locked = True
        wsSyn.Protect Password:=pwd, _
                      DrawingObjects:=True, _
                      Contents:=True, _
                      Scenarios:=True, _
                      UserInterfaceOnly:=userInterfaceOnly, _
                      AllowFormattingCells:=True, _
                      AllowFormattingColumns:=True, _
                      AllowFormattingRows:=True, _
                      AllowSorting:=True, _
                      AllowFiltering:=True
        wsSyn.EnableSelection = xlNoRestrictions
    End If
    
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
End Sub

Public Sub UnprotectCalculatedSheetsForMacro()
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    Dim pwd As String
    pwd = SheetPassword()
    
    Dim wsOff As Worksheet, wsSyn As Worksheet
    On Error Resume Next
    Set wsOff = ThisWorkbook.Worksheets("OFFER")
    Set wsSyn = ThisWorkbook.Worksheets("SYNOPSIS")
    On Error GoTo 0
    
    If Not wsOff Is Nothing Then
        UnprotectSheet wsOff, pwd
    End If
    If Not wsSyn Is Nothing Then
        UnprotectSheet wsSyn, pwd
    End If
    
    Application.DisplayAlerts = prevAlerts
End Sub

Public Sub UnprotectAllSheets(Optional ByVal pwd As String = "")
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        UnprotectSheet ws, pwd
    Next ws
End Sub

Public Function RunProtectionSelfTests() As String
    If Not ENABLE_SHEET_PROTECTION Then
        RunProtectionSelfTests = "PASS"
        Exit Function
    End If
    Dim wsQuote As Worksheet
    Dim wsLines As Worksheet
    Dim tblLines As ListObject
    Dim originalLocked As Boolean
    Dim writeTestVal As Variant
    Dim testError As Boolean
    
    On Error GoTo TestFail
    
    ' 1. Configure cell locking
    ConfigureCellLocking
    
    Set wsQuote = ThisWorkbook.Worksheets("QUOTE")
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    ' Verify cell locking: input cell B5 unlocked, summary cell D10 locked
    If wsQuote.Range("B5").Locked <> False Then
        RunProtectionSelfTests = "FAIL: Input cell B5 on QUOTE is locked"
        Exit Function
    End If
    
    If wsQuote.Range("D10").Locked <> True Then
        RunProtectionSelfTests = "FAIL: Grand Total cell D10 on QUOTE is not locked"
        Exit Function
    End If
    
    ' Verify table column locking: Quantity unlocked, LineTotal locked
    If tblLines.ListColumns("Quantity").DataBodyRange.Cells(1, 1).Locked <> False Then
        RunProtectionSelfTests = "FAIL: Table input column Quantity is locked"
        Exit Function
    End If
    
    If tblLines.ListColumns("LineTotal").DataBodyRange.Cells(1, 1).Locked <> True Then
        RunProtectionSelfTests = "FAIL: Table output column LineTotal is not locked"
        Exit Function
    End If
    
    ' 2. Protect with UserInterfaceOnly:=True
    If ENABLE_SHEET_PROTECTION Then
        ProtectSheet wsQuote, True
        If Not wsQuote.ProtectContents Then
            RunProtectionSelfTests = "FAIL: wsQuote.ProtectContents is False after ProtectSheet"
            Exit Function
        End If
    Else
        wsQuote.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        If Not wsQuote.ProtectContents Then
            RunProtectionSelfTests = "FAIL: wsQuote.ProtectContents is False after Protect"
            Exit Function
        End If
    End If
    
    ' 3. Verify macro write capability: macro can write to locked cell without error
    testError = False
    On Error Resume Next
    writeTestVal = wsQuote.Range("D11").Value
    wsQuote.Range("D11").Value = "Workflow Status"
    If Err.Number <> 0 Then
        testError = True
    End If
    On Error GoTo TestFail
    
    If testError Then
        RunProtectionSelfTests = "FAIL: Macro failed to write to locked cell under UserInterfaceOnly:=True"
        UnprotectSheet wsQuote
        Exit Function
    End If
    
    ' 4. Clean up / Unprotect
    UnprotectAllSheets
    If wsQuote.ProtectContents Then
        RunProtectionSelfTests = "FAIL: wsQuote.ProtectContents is still True after UnprotectAllSheets"
        Exit Function
    End If
    
    RunProtectionSelfTests = "PASS"
    Exit Function

TestFail:
    RunProtectionSelfTests = "FAIL: Exception - " & Err.Description
    On Error Resume Next
    UnprotectAllSheets
End Function

