# XLSX to XLSM Conversion Procedure

The development workbook is intentionally created first as a valid
`DEV/Mega_Pipes_Quotation_DEV.xlsx`. It must be opened and converted by
desktop Microsoft Excel before VBA is imported.

1. Open `DEV\Mega_Pipes_Quotation_DEV.xlsx` in desktop Microsoft Excel.
2. Use **File -> Save As**.
3. Select **Excel Macro-Enabled Workbook (*.xlsm)**.
4. Save as `DEV\Mega_Pipes_Quotation_DEV.xlsm`.
5. Close Excel.
6. Reopen the newly saved XLSM.
7. Press **ALT+F11**.
8. Import the VBA modules from `VBA\Modules` using **File -> Import File**.
9. Save the workbook.

The XLSM conversion must be performed by Microsoft Excel itself before VBA is
imported. Do not change an extension or manually edit OOXML content types to
claim that an XLSX package is an XLSM.

The current workbook has not completed this manual conversion. XLSM validity is
therefore pending manual Excel Save As and reopen verification.
