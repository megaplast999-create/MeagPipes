Attribute VB_Name = "modNavigation"
Option Explicit

Public Sub GoToStart()
    ThisWorkbook.Worksheets("START").Activate
End Sub

Public Sub EnterSystem()
    ThisWorkbook.Worksheets("QUOTE").Activate
End Sub

Public Sub GoToQuote()
    ThisWorkbook.Worksheets("QUOTE").Activate
End Sub

Public Sub GoToCosting()
    ThisWorkbook.Worksheets("COSTING").Activate
End Sub

Public Sub GoToOffer()
    ThisWorkbook.Worksheets("OFFER").Activate
End Sub
