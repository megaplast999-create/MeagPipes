# VBA Import and Embedding Status

The development workbook currently does not contain a verified
`xl/vbaProject.bin`. It must not be faked by renaming an XLSX package.

Importable source files are stored under:

- `VBA/Modules/modMain.bas`
- `VBA/Modules/modNavigation.bas`
- `VBA/Modules/modGeometry.bas`
- `VBA/Modules/modCalculation.bas`
- `VBA/Modules/modValidation.bas`

To embed these modules, use the Excel VBA editor in a trusted local copy:

1. Open `DEV/Mega_Pipes_Quotation_DEV.xlsm`.
2. Open the VBA editor with `Alt+F11`.
3. Use **File > Import File** for each `.bas` module.
4. Compile the VBA project.
5. Save the workbook as a real macro-enabled workbook.
6. Reopen it and verify that `xl/vbaProject.bin` is present in the package.

Embedding is now **VERIFIED for the DEV P1 checkpoint**. Excel 16.0 was
opened through pywin32 after the approved Trust Center changes; VBOM access
reported 23 components, the drifted source modules were imported, the VBE
compile command was executed, and the header/quote-line smoke test passed.
See `TESTS/excel_runtime_p1_20261008.txt`. This does not claim that later
calculation, ORS, export, or multi-user phases are complete.
