<#
.SYNOPSIS
Exports Active Directory mail-enabled groups and their members to Excel.

.DESCRIPTION
Retrieves the following information from Active Directory:
- User mail addresses (mail, UPN, proxyAddresses)
- Mail-enabled groups (Distribution Groups and Mail-Enabled Security Groups)
- Group members including nested memberships

The collected data is exported into a single Excel file with multiple worksheets.

.OUTPUTS
Excel file (.xlsx)

.REQUIREMENTS
- RSAT / ActiveDirectory PowerShell module
- ImportExcel PowerShell module (Microsoft Excel is not required)

.AUTHOR
Ceyhun Yıldız

.DATE
2026-01-20
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# =========================
# AYARLAR (Sadece burayı değiştirmen yeterli)
# =========================
$OutputDir  = "C:\AD_PS_Ops\Reports"
$FileName   = "AD_Mail_Groups_And_Members_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$OutputPath = Join-Path $OutputDir $FileName

# Excel görünümü
$TableStyle    = "Medium2"
$FreezeTopRow  = $true
$AutoSize      = $true
$BoldTopRow    = $true

# Mail-enabled grup tespiti (LDAP)
$GroupLdapFilter = "(&(objectCategory=group)(|(mail=*)(proxyAddresses=*)(mailNickname=*)))"

# =========================
# ÖN KONTROLLER
# =========================
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "ActiveDirectory modülü bulunamadı. Lütfen RSAT / AD PowerShell modülünü kurun."
}
Import-Module ActiveDirectory -ErrorAction Stop

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    try {
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    } catch {
        throw "ImportExcel modülü kurulamadı. PowerShell'i yönetici açın veya manuel kurun. Hata: $($_.Exception.Message)"
    }
}
Import-Module ImportExcel -ErrorAction Stop

function TS {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return "" }
    return $Text.Trim()
}

# =========================
# 1) KULLANICILAR (mail)
# =========================
Write-Host "1/3 Kullanıcılar çekiliyor..." -ForegroundColor Cyan

$Kullanicilar = Get-ADUser -Filter * -Properties mail, proxyAddresses, userPrincipalName, displayName, sAMAccountName, enabled |
    Select-Object `
        @{n="DisplayName";e={TS $_.displayName}},
        @{n="SamAccountName";e={TS $_.sAMAccountName}},
        @{n="UPN";e={TS $_.userPrincipalName}},
        @{n="Mail";e={TS $_.mail}},
        @{n="ProxyAddresses";e={ if ($_.proxyAddresses) { ($_.proxyAddresses | ForEach-Object { TS $_ }) -join "; " } else { "" } }},
        @{n="Enabled";e={$_.Enabled}},
        @{n="DN";e={TS $_.DistinguishedName}}

# =========================
# 2) MAIL GRUPLARI
# =========================
Write-Host "2/3 Mail grupları çekiliyor..." -ForegroundColor Cyan

$GruplarRaw = Get-ADGroup -LDAPFilter $GroupLdapFilter -Properties mail, proxyAddresses, mailNickname, groupCategory, groupScope, displayName, managedBy

$Gruplar = $GruplarRaw |
    Select-Object `
        @{n="GrupAdi";e={TS $_.displayName}},
        @{n="SamAccountName";e={TS $_.sAMAccountName}},
        @{n="GrupMail";e={TS $_.mail}},
        @{n="ProxyAddresses";e={ if ($_.proxyAddresses) { ($_.proxyAddresses | ForEach-Object { TS $_ }) -join "; " } else { "" } }},
        @{n="MailNickname";e={TS $_.mailNickname}},
        @{n="GroupCategory";e={TS $_.groupCategory}},
        @{n="GroupScope";e={TS $_.groupScope}},
        @{n="ManagedByDN";e={TS $_.managedBy}},
        @{n="DN";e={TS $_.DistinguishedName}} |
    Sort-Object GrupAdi

Write-Host ("Bulunan mail-enabled grup sayısı: {0}" -f $Gruplar.Count) -ForegroundColor Yellow

# =========================
# 3) GRUP ÜYELERİ (nested)
# =========================
Write-Host "3/3 Grup üyeleri çıkarılıyor (nested/recursive)..." -ForegroundColor Cyan

$GrupUyeleri = foreach ($g in $Gruplar) {
    $rawMembers = $null
    try {
        $rawMembers = Get-ADGroupMember -Identity $g.DN -Recursive -ErrorAction Stop
    } catch {
        [pscustomobject]@{
            GrupAdi           = $g.GrupAdi
            GrupMail          = $g.GrupMail
            UyeTipi           = "HATA"
            UyeDisplayName    = "Üyeler okunamadı: $($_.Exception.Message)"
            UyeSamAccountName = ""
            UyeUPN            = ""
            UyeMail           = ""
            UyeDN             = $g.DN
        }
        continue
    }

    foreach ($m in $rawMembers) {
        $uyeTip = switch ($m.objectClass) {
            "user"  { "Kullanıcı" }
            "group" { "Grup" }
            default { TS $m.objectClass }
        }

        $uyeDisplay = TS $m.Name
        $uyeSam     = TS $m.SamAccountName
        $uyeDN      = TS $m.DistinguishedName
        $uyeUPN     = ""
        $uyeMail    = ""

        if ($m.objectClass -eq "user") {
            try {
                $u = Get-ADUser -Identity $m.DistinguishedName -Properties mail, userPrincipalName, displayName, sAMAccountName -ErrorAction Stop
                $uyeDisplay = TS $u.displayName
                $uyeSam     = TS $u.sAMAccountName
                $uyeUPN     = TS $u.userPrincipalName
                $uyeMail    = TS $u.mail
            } catch { }
        }
        elseif ($m.objectClass -eq "group") {
            try {
                $gg = Get-ADGroup -Identity $m.DistinguishedName -Properties mail, displayName, sAMAccountName -ErrorAction Stop
                $uyeDisplay = TS $gg.displayName
                $uyeSam     = TS $gg.sAMAccountName
                $uyeMail    = TS $gg.mail
            } catch { }
        }

        [pscustomobject]@{
            GrupAdi           = $g.GrupAdi
            GrupMail          = $g.GrupMail
            UyeTipi           = $uyeTip
            UyeDisplayName    = $uyeDisplay
            UyeSamAccountName = $uyeSam
            UyeUPN            = $uyeUPN
            UyeMail           = $uyeMail
            UyeDN             = $uyeDN
        }
    }
}

Write-Host ("Toplam üye satırı: {0}" -f ($GrupUyeleri | Measure-Object).Count) -ForegroundColor Yellow

# =========================
# EXCEL'E AKTAR
# =========================
Write-Host "Excel'e aktarılıyor..." -ForegroundColor Cyan

if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Force
}

$exportParams = @{
    Path         = $OutputPath
    AutoSize     = $AutoSize
    FreezeTopRow = $FreezeTopRow
    BoldTopRow   = $BoldTopRow
    TableStyle   = $TableStyle
}

$Kullanicilar | Export-Excel @exportParams -WorksheetName "Kullanicilar"
$Gruplar      | Export-Excel @exportParams -WorksheetName "MailGruplari"
$GrupUyeleri  | Export-Excel @exportParams -WorksheetName "GrupUyeleri"

Write-Host "Tamamlandı." -ForegroundColor Green
Write-Host "Çıktı: $OutputPath" -ForegroundColor Green

Invoke-Item $OutputPath
