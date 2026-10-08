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

Embedding is currently **BLOCKED/PENDING** because Excel COM automation is
unreliable in this environment and no safe verified VBA project was available.
The source files are not a claim that the workbook macros are already active.
