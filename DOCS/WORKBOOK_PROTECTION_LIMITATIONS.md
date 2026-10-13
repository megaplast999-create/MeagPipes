# Mega Pipes Quotation System — Workbook Protection Architecture & Limitations

## 1. Overview
Phase 23 implements sheet and cell protection to enforce data integrity:
- **Locked cells** (Calculated fields, summary card, master tables, distance cache, audit log): Prevent manual tampering or accidental overwriting by users.
- **Unlocked cells** (Customer dropdown, project name, delivery PIN, pipe line inputs): Remain editable for normal quotation workflow.
- **Macro Writes (`UserInterfaceOnly:=True`)**: Enables VBA automation to write to locked cells (summary card totals, calculated OD/ID, allocated quote numbers) without requiring continuous unprotect/protect cycles.

## 2. Technical Limitations & Security Realities

> [!WARNING]
> Excel worksheet protection is an interface guard designed to prevent accidental user errors and guide data entry. It is **not** a cryptographic security boundary against determined technical adversaries.

1. **UserInterfaceOnly Non-Persistence**:
   - Excel does not persist the `UserInterfaceOnly:=True` argument when a workbook is saved and closed.
   - Upon reopening, Excel defaults to standard sheet protection where macros cannot write to locked cells unless `ProtectAllSheets UserInterfaceOnly:=True` is re-applied during `Workbook_Open` or before automation procedures.
2. **Password Storage & Governance**:
   - Sheet and admin passwords are never hardcoded in VBA source code.
   - Sourced dynamically via environment variables:
     - `MEGA_SHEET_PWD`: Sheet protection password.
     - `MEGA_ADMIN_PWD`: Admin master synchronization and approval password.
   - If `MEGA_SHEET_PWD` is unset in the environment, protection is applied with an empty password string (still blocking UI edits, but permitting unprotect without prompt).
3. **Third-Party OpenXML / Python Editors**:
   - Worksheet protection hashes in OpenXML files (`sheetProtection` tags in `.xml`) can be removed by external XML editing tools or scripts. True access control for confidential files relies on filesystem/share permissions.
4. **Human Action Required**:
   - Production deployment requires configuring `MEGA_SHEET_PWD` and `MEGA_ADMIN_PWD` on all authorized machines. Cataloged in `TESTS/NEEDS_HUMAN.md`.

