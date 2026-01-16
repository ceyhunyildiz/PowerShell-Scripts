<#
.SYNOPSIS
Exports Active Directory user attributes to an Excel-compatible CSV file.

.DESCRIPTION
Retrieves a wide range of Active Directory user attributes and exports
the collected data to a CSV file (default delimiter ';') for reporting and analysis.

Optionally opens the exported file after completion.

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
$OutputPath      = "C:\AD_PS_Ops\Reports\AD_Kullanicilar_Attribute.csv"
$Delimiter       = ';'

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
Import-Module ActiveDirectory

# Çıkış dizinini oluştur
New-Item -ItemType Directory -Path (Split-Path $OutputPath) -Force | Out-Null

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
    SamAccountName,Description,POBox,ObjectClass,DisplayName,GivenName,Surname,Name,CanonicalName,
    DistinguishedName,Company,Department,Title,Manager,EmailAddress,MobilePhone,OfficePhone,Office,
    StreetAddress,City,State,Created,LastLogonDate,LastBadPasswordAttempt

$users | Export-Csv -Path $OutputPath -NoTypeInformation -Delimiter $Delimiter -Encoding UTF8

Write-Host "Bitti. Dosya: $OutputPath"

if ($OpenAfterExport) {
    Invoke-Item -Path $OutputPath
}
