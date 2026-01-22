# PowerShell Scripts – Active Directory & DFS Automation

This repository contains PowerShell scripts focused on **Active Directory reporting** and **DFS-based file operations**, created for daily system administration and IT support tasks.

The scripts follow these principles:

- **Run-and-work** (no complex parameters required)
- Clear structure and documentation
- Enterprise-oriented output (**Excel, CSV, HTML, Logs**)
- Easy customization via a simple **SETTINGS / AYARLAR** section

---

## 📌 Requirements

- **Windows PowerShell 5.1**
- **RSAT – Active Directory module** (for AD scripts)
- **ImportExcel PowerShell module** (for `.xlsx` outputs)
  - Microsoft Excel **is not required** when using ImportExcel
- Appropriate permissions on **Active Directory** and **file shares**
- (Optional) Network access and rights for DFS/Share operations

---

## 📂 Scripts Overview

### 🔹 Copy-PSTFilesFromDFS.ps1
Copies Outlook PST files from a DFS-based file system to a local directory using **robocopy** for performance and reliability.

**Key features**
- Recursive search for `.pst` files
- Multi-threaded copy support (robocopy)
- Automatic log generation
- Designed for large PST collections

**Outputs**
- Log file(s) under the output folder (depending on settings)

**Typical use cases**
- PST collection before migrations
- Cleanup or audit operations
- DFS-based file analysis

---

### 🔹 Export-ADUsersToExcel.ps1
Exports a wide range of Active Directory **user attributes** to an Excel (`.xlsx`) file.

**Key features**
- Comprehensive AD attribute export
- Excel formatting (filters, freeze panes, auto-fit)
- Timestamped output
- Excel file automatically opens after export

**Outputs**
- `.xlsx` report (formatted)

**Typical use cases**
- User inventory reports
- HR and audit reporting
- Bulk AD analysis

---

### 🔹 Export-ADMailGroupsAndMembers.ps1
Exports **mail-enabled AD groups** and their members, including **nested memberships**, to Excel for reporting/audit purposes.

**Key features**
- Identifies mail-enabled groups (mail/proxyAddresses/mailNickname logic)
- Exports group details + membership
- Nested (recursive) member resolution
- Excel output with enterprise formatting

**Outputs**
- `.xlsx` report (typically multi-sheet depending on script structure)

**Typical use cases**
- Mail group inventory & audits
- Distribution/security group membership reviews
- Access and permission verification (mail-enabled structures)

---

### 🔹 Export-AD_MultiGroupMembers.ps1
Prompts for **one or more AD groups (comma-separated)** and exports **all user members (nested included)** into a single Excel report.

**Key features**
- Prompts for group name(s): `GroupA, GroupB, GroupC`
- Nested (recursive) membership support
- De-duplicates users across multiple groups
- Adds a column showing **which groups** each user belongs to (e.g., `Üye Olduğu Gruplar`)
- Excel output auto-opens after export

**Outputs**
- `.xlsx` report (formatted)

**Typical use cases**
- Multi-group access reviews
- Role-based membership audits
- Comparing membership across multiple teams/roles

---

### 🔹 Get-ADFLGroupMembers.ps1
Retrieves members of AD groups whose names end with **"FL"**, including nested group memberships.

**Key features**
- Filters groups by naming convention (suffix-based)
- Recursive membership extraction
- Generates both machine-readable and quick-view outputs
- CSV auto-opens in Excel

**Outputs**
- `.csv` report (Excel compatible)
- `.html` report for quick viewing

**Typical use cases**
- Role-based group analysis
- Nested group structure auditing
- Access and permission reviews

---

### 🔹 Get-ADRecentUsers.ps1
Exports Active Directory users created within the **last 7 days** to an Excel file.

**Key features**
- Filters users by creation date
- Excel output with formatting
- Automatically opens the generated file
- Useful for weekly change tracking

**Outputs**
- `.xlsx` report (formatted)

**Typical use cases**
- Weekly user creation audits
- New user onboarding checks
- Change tracking in Active Directory

---

## 📁 Output Structure

By default, scripts generate their output under a dedicated output directory.
Output files may include:

- Excel reports (`.xlsx`)
- CSV reports (`.csv`)
- HTML reports (`.html`)
- Log files (e.g., robocopy logs)

The output directory can be customized in the **SETTINGS / AYARLAR** section at the beginning of each script.

---

## 🔐 Notes

- This repository does **not** contain credentials.
- Environment-specific values (domain names, OU paths, file share paths) should be adjusted in each script’s **settings** section.
- Always test scripts in a **non-production** environment before running in production.

---
