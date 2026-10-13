# Mega Pipes Quotation System — Human Action Required Items

## 1. ORS_API_KEY (Phase 9 Live Check)
- **Status:** `NEEDS_HUMAN`
- **Reason:** Environment variable `ORS_API_KEY` is not configured in the current Windows environment.
- **Action Required:**
  Set the OpenRouteService API key in the Windows environment:
  ```powershell
  [System.Environment]::SetEnvironmentVariable("ORS_API_KEY", "your_ors_api_key_here", "User")
  ```
- **Fallback Activated:**
  - Mocked response tests and cache tests: `PASS`
  - Seeded distance cache from Unipro reference workbook: Active (`782 km` for Jabalpur, `1821 km` for Khopoli to Bagodar 825322).
  - Manual distance fallback: Fully operational.
- **Safety Guarantee:** No fake keys used; live check skipped safely without blocking downstream phases.

## 2. Passwords: MEGA_SHEET_PWD & MEGA_ADMIN_PWD (Phases 21, 23)
- **Status:** `NEEDS_HUMAN`
- **Reason:** Passwords are not hardcoded in source code and require configuration in the Windows environment for production deployments.
- **Action Required:**
  Set production passwords in the Windows environment:
  ```powershell
  [System.Environment]::SetEnvironmentVariable("MEGA_SHEET_PWD", "your_sheet_protection_password", "User")
  [System.Environment]::SetEnvironmentVariable("MEGA_ADMIN_PWD", "your_admin_approval_password", "User")
  ```
- **Fallback Activated:**
  - When unset, development macros proceed with empty-string password validation.
  - Full structural and cell locking protection remains verified.
