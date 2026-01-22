<#
.SYNOPSIS
Exports Active Directory user attributes to an Excel (.xlsx) file.

.DESCRIPTION
Retrieves a wide range of Active Directory user attributes and exports
the collected data to an Excel file for reporting and analysis.

The generated Excel file is automatically opened after export.

.AUTHOR
Ceyhun Yıldız

.DATE
2026-01-16
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =========================
# AYARLAR (Sadece burayı değiştirmen yeterli)
# =========================
$OutputDir       = "C:\AD_PS_Ops\Reports"
$FileName        = "AD_Kullanicilar_Attribute_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$OutputPath      = Join-Path $OutputDir $FileName

# Excel otomatik açılsın mı?
$OpenAfterExport = $true

# İstersen belirli bir OU’dan çek (boş bırakırsan tüm AD)
# Örn: "OU=Users,DC=domain,DC=local"
$SearchBase      = ""

# =========================
# KOD
# =========================

# ActiveDirectory modülü kontrol
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "ActiveDirectory modülü bulunamadı. RSAT (Active Directory module) kurulu olmalı."
}
Import-Module ActiveDirectory -ErrorAction Stop

# Çıkış dizinini oluştur
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

# Çekilecek attribute listesi
$props = @(
    "CanonicalName","City","Company","Created","Department","Description","DisplayName","DistinguishedName",
    "EmailAddress","GivenName","LastBadPasswordAttempt","LastLogonDate","Manager","MobilePhone","Office",
    "OfficePhone","POBox","SamAccountName","State","StreetAddress","Surname","Title","ObjectClass","Name"
)

$adParams = @{
    Filter     = '*'
    Properties = $props
}
if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
    $adParams.SearchBase = $SearchBase
}

$users = Get-ADUser @adParams | Select-Object `
    SamAccountName,DisplayName,GivenName,Surname,Name,Description,POBox,ObjectClass,
    Company,Department,Title,Manager,EmailAddress,MobilePhone,OfficePhone,Office,
    StreetAddress,City,State,CanonicalName,DistinguishedName,
    Created,LastLogonDate,LastBadPasswordAttempt

# =========================
# EXCEL (COM)
# =========================
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    throw "Excel COM objesi oluşturulamadı. Bu makinede Microsoft Excel yüklü olmayabilir."
}

$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)
$sheet.Name = "AD_Users"

try {
    # Başlıklar
    $headers = $users[0].PSObject.Properties.Name
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $sheet.Cells.Item(1, $c + 1).Value2 = $headers[$c]
        $sheet.Cells.Item(1, $c + 1).Font.Bold = $true
    }

    # Veriler
    $row = 2
    foreach ($u in $users) {
        $col = 1
        foreach ($h in $headers) {
            $sheet.Cells.Item($row, $col).Value2 = $u.$h
            $col++
        }
        $row++
    }

    # Biçimlendirme
    $used = $sheet.UsedRange
    $used.EntireColumn.AutoFit() | Out-Null
    $sheet.Range("A1").AutoFilter() | Out-Null
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    # Kaydet / Kapat
    $workbook.SaveAs($OutputPath)
    $workbook.Close($true)
    $excel.Quit()
}
finally {
    # COM temizliği
    if ($sheet)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)    | Out-Null }
    if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)    | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Excel'i otomatik aç
if ($OpenAfterExport -and (Test-Path $OutputPath)) {
    Start-Process $OutputPath
}

Write-Host "Bitti. Excel oluşturuldu ve açıldı:" -ForegroundColor Green
Write-Host $OutputPath
