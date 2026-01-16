# PowerShell Scripts – Active Directory & DFS Automation

This repository contains PowerShell scripts focused on **Active Directory reporting**
and **DFS-based file operations**, created for daily system administration
and IT support tasks.

The scripts are designed with the following principles:
- Run-and-work (no complex parameters required)
- Clear structure and documentation
- Enterprise-oriented output (Excel, CSV, HTML)
- Easy customization via a simple settings section

---

## 📌 Requirements

- Windows PowerShell 5.1
- RSAT – Active Directory module
- Microsoft Excel (required for `.xlsx` output scripts)
- Appropriate permissions on Active Directory and file shares

---

## 📂 Scripts Overview

### 🔹 Copy-PSTFilesFromDFS.ps1
Copies Outlook PST files from a DFS-based file system to a local directory
using **robocopy** for performance and reliability.

**Key features**
- Recursive search for `.pst` files
- Multi-threaded copy
- Automatic log generation
- Designed for large PST collections

**Typical use cases**
- PST collection before migrations
- Cleanup or audit operations
- DFS-based file analysis

---

### 🔹 Export-ADUsersToExcel.ps1
Exports a wide range of Active Directory user attributes to an **Excel (.xlsx)** file.

**Key features**
- Comprehensive AD attribute export
- Excel formatting (filters, freeze panes, auto-fit)
- Timestamped output
- Excel file automatically opens after export

**Typical use cases**
- User inventory reports
- HR and audit reporting
- Bulk AD analysis

---

### 🔹 Get-ADFLGroupMembers.ps1
Retrieves members of Active Directory groups whose names end with **"FL"**,
including **nested group memberships**.

**Outputs**
- CSV report (Excel compatible)
- HTML report for quick viewing
- CSV file automatically opens in Excel

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

**Typical use cases**
- Weekly user creation audits
- New user onboarding checks
- Change tracking in Active Directory

---

## 📁 Output Structure

By default, all scripts generate their output under the following directory:
Output files may include Excel reports (`.xlsx`), CSV files (`.csv`),
HTML reports, and log files depending on the script.
The output directory can be customized in the **SETTINGS / AYARLAR**
section at the beginning of each script.

---

## 🔐 Notes

This repository does not contain credentials.
Environment-specific values (such as domain names or file paths)
should be adjusted in the settings section of each script.
Always test scripts in a non-production environment before use.
