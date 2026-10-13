Attribute VB_Name = "modProtection"
Option Explicit

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
    Dim pwd As String
    pwd = Environ$("MEGA_SHEET_PWD")
    SheetPassword = pwd
End Function

Public Sub ConfigureCellLocking()
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
    wsQuote.Cells.Locked = True
    
    ' Unlock editable header inputs: Customer (B5) through Price Basis (B17)
    wsQuote.Range("B5:B17").Locked = False
    
    ' Unlock editable product item inputs (B21:B31)
    wsQuote.Range("B21:B31").Locked = False
    
    ' Keep Enquiry Number (B4) and Summary card (D3:E14) locked
    wsQuote.Range("B4").Locked = True
    wsQuote.Range("D3:E14").Locked = True
    
    ' Unlock feedback banner cell
    wsQuote.Range("A33").Locked = False
    
    ' 2. QUOTE_LINES Sheet: Unlock user input columns in tblQuoteLines
    Set wsLines = ThisWorkbook.Worksheets("QUOTE_LINES")
    Set tblLines = wsLines.ListObjects("tblQuoteLines")
    
    wsLines.Cells.Locked = True
    
    ' Unlock user data-entry columns
    UnlockColumn tblLines, "Product"
    UnlockColumn tblLines, "Family"
    UnlockColumn tblLines, "Basis"
    UnlockColumn tblLines, "SizeMM"
    UnlockColumn tblLines, "WT_MM"
    UnlockColumn tblLines, "Grade"
    UnlockColumn tblLines, "Quantity"
    UnlockColumn tblLines, "QuantityUnit"
    UnlockColumn tblLines, "LengthM"
    UnlockColumn tblLines, "SupplyingPlant"
    UnlockColumn tblLines, "ExternalCoating"
    UnlockColumn tblLines, "InternalCoating"
    
    ' Lock calculation and status columns
    LockColumn tblLines, "LineID"
    LockColumn tblLines, "PipeCount"
    LockColumn tblLines, "TotalMetres"
    LockColumn tblLines, "TotalMT"
    LockColumn tblLines, "CalculatedODMM"
    LockColumn tblLines, "CalculatedIDMM"
    LockColumn tblLines, "UnitRateINR"
    LockColumn tblLines, "LineTotal"
    LockColumn tblLines, "Status"
    
    ' 3. MASTERS Sheet: Unlock cells so dynamic cascading dropdown helpers can write
    On Error Resume Next
    ThisWorkbook.Worksheets("MASTERS").Cells.Locked = False
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
    Dim pwd As String
    pwd = SheetPassword()
    
    On Error Resume Next
    ws.Protect Password:=pwd, _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True, _
               UserInterfaceOnly:=userInterfaceOnly, _
               AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, _
               AllowFormattingRows:=True, _
               AllowSorting:=True, _
               AllowFiltering:=True
    On Error GoTo 0
End Sub

Public Sub UnprotectSheet(ByVal ws As Worksheet, Optional ByVal pwd As String = "")
    Dim effectivePwd As String
    effectivePwd = pwd
    If Len(effectivePwd) = 0 Then effectivePwd = SheetPassword()
    
    On Error Resume Next
    ws.Unprotect Password:=effectivePwd
    On Error GoTo 0
End Sub

Public Sub ProtectAllSheets(Optional ByVal userInterfaceOnly As Boolean = True)
    Dim ws As Worksheet
    ConfigureCellLocking
    For Each ws In ThisWorkbook.Worksheets
        ProtectSheet ws, userInterfaceOnly
    Next ws
End Sub

Public Sub UnprotectAllSheets(Optional ByVal pwd As String = "")
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        UnprotectSheet ws, pwd
    Next ws
End Sub

Public Function RunProtectionSelfTests() As String
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
    ProtectSheet wsQuote, True
    If Not wsQuote.ProtectContents Then
        RunProtectionSelfTests = "FAIL: wsQuote.ProtectContents is False after ProtectSheet"
        Exit Function
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

