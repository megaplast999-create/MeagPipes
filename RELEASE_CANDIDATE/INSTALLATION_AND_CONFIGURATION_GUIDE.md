# Mega Pipes Quotation System — Installation & Configuration Guide

**Version:** v1.0.0-RC  
**Target Audience:** Enterprise IT Administrators, System Engineers, Commercial Operations Leads  
**Applicable Platform:** Microsoft Windows 10/11 / Windows Server, Microsoft Office 2016 / 2019 / 2021 / Office 365

---

## 1. System Requirements & Prerequisites

### 1.1 Workstation Requirements
- **Operating System:** Windows 10 (Build 19041+) or Windows 11 (64-bit recommended).
- **Office Suite:** Microsoft Excel 2016, 2019, 2021, or Microsoft 365 Desktop Application.
- **VBA Support:** Visual Basic for Applications (VBA 7.1) must be installed.
- **Runtime Libraries:** Microsoft Scripting Runtime (`scrrun.dll`), Microsoft XML v6.0 (`msxml6.dll`) enabled.

### 1.2 Network Infrastructure
- **Network Protocol:** SMB 2.1 or SMB 3.0+ file sharing.
- **Network Share:** Dedicated enterprise file server UNC path (e.g., `\\fileserver\MegaPipes\QuotationShared\`).
- **Permissions:**
  - Estimators / Sales Users: **Read, Write, Modify** on the `Quotes\`, `Sequence\`, and `Masters\` folders.
  - IT Administrators / Commercial Managers: **Full Control**.

---

## 2. Directory Structure Setup

Create the standardized production shared folder hierarchy on the corporate network file server:

```text
\\fileserver\MegaPipes\QuotationShared\
├── Quotes\
│   ├── 2026-27\
│   └── 2027-28\
├── Sequence\
│   ├── sequence.txt
│   └── sequence.lock
├── Masters\
│   └── Masters_Central.xlsx
└── Backups\
    └── Daily\
```

### Initial File Deployment
1. **Sequence File:** Create an empty text file `Sequence\sequence.txt` containing the initial starting counter (e.g., `594` or desired next integer) if seeding from a legacy system.
2. **Central Masters:** Place the approved, finalized master database workbook into `Masters\Masters_Central.xlsx`.
3. **Application Distribution:** Distribute `Mega_Pipes_Quotation_v1.0.0_RC.xlsm` to local user workstations or publish it in a read-only shared software distribution folder.

---

## 3. Microsoft Excel Trust Center Configuration

To ensure uninterrupted execution of macros without security warning bars or disabled script execution:

### 3.1 Adding a Trusted Location
1. Open Microsoft Excel and navigate to **File** -> **Options** -> **Trust Center** -> **Trust Center Settings...**.
2. Select **Trusted Locations** on the left menu.
3. Click **Add new location...**.
4. Enter the path to the network share (e.g., `\\fileserver\MegaPipes\QuotationShared`) or the local workstation deployment folder (e.g., `C:\MegaPipes\App`).
5. If using a network share, check the box: **"Allow Trusted Locations on my network (not recommended)"** and check **"Subfolders of this location are also trusted"**.
6. Click **OK** to save.

### 3.2 Macro Settings Configuration
1. In the **Trust Center Settings** dialog, select **Macro Settings**.
2. Select **"Disable VBA macros with notification"** (or **"Enable VBA macros"** if governed by corporate Active Directory Group Policy).
3. Ensure **"Trust access to the VBA project object model"** is checked for administrative diagnostic stations.
4. Click **OK** and close Excel options.

---

## 4. Environment Variables Configuration

The Mega Pipes Quotation System utilizes Windows environment variables to separate operational credentials and network paths from workbook code.

### 4.1 Production Environment Variables

| Variable Name | Purpose | Production Example | Default Fallback |
|---|---|---|---|
| `ORS_API_KEY` | OpenRouteService API key for real-time routing distance matrix queries | `5b3ce3597851110001cf6248...` | Built-in distance cache (`tblDistanceCache`) + Manual distance input |
| `MEGA_SHEET_PWD` | Password for worksheet structure and cell protection | `MegaProd$heet2026` | Empty password (DEV mode) |
| `MEGA_ADMIN_PWD` | Password for commercial quote approval and master data synchronization | `AdminApprove#2026` | Empty password (DEV mode) |
| `MEGA_SHARED_FOLDER` | Corporate UNC share root path | `\\fileserver\MegaPipes\QuotationShared` | Local workstation `Quotes\` folder |

### 4.2 Configuration Commands (PowerShell)
To set system-wide or user-level environment variables across user workstations, execute via PowerShell (run as Administrator for Machine scope):

```powershell
# User-level configuration (per workstation login)
[System.Environment]::SetEnvironmentVariable("ORS_API_KEY", "your_production_ors_key_here", "User")
[System.Environment]::SetEnvironmentVariable("MEGA_SHEET_PWD", "your_secure_sheet_password", "User")
[System.Environment]::SetEnvironmentVariable("MEGA_ADMIN_PWD", "your_secure_admin_password", "User")
[System.Environment]::SetEnvironmentVariable("MEGA_SHARED_FOLDER", "\\fileserver\MegaPipes\QuotationShared", "User")
```

---

## 5. Post-Installation Verification & Acceptance Testing

Before granting sales estimators access, execute the following pre-flight verification steps on at least one representative workstation:

1. **Launch Workbook:** Open `Mega_Pipes_Quotation_v1.0.0_RC.xlsm`. Verify that no compile error or yellow security prompt appears.
2. **Execute Unified Self-Test Suite:**
   - Press `Alt + F8`, select `RunAllSelfTests`, and click **Run** (or click the **Self Test** button on the `START` sheet).
   - Ensure the execution completes in under 10 seconds.
   - Review output in `TESTS\unified_selftest_*.txt` or message dialog: confirm 20 automated tests report **PASS** and only external credentials (`ORS_API_KEY`, passwords) report `NEEDS_HUMAN` if not configured.
3. **Verify Sequence Lock:**
   - On the `QUOTE` sheet, enter customer name `Test Installation`.
   - Click **Generate Enquiry No.**.
   - Verify that an official enquiry number (e.g., `MKTGENQ26-27-001`) is generated and recorded in the sequence tracking system.
4. **Export Sanity Test:**
   - Click **Export Offer (PDF)**. Verify that a clean 1-page PDF is generated without error.
   - Click **Export Customer (Excel)**. Open the generated file and confirm that only `OFFER` and `Synopsis` sheets are present, and zero calculation formulas exist.
