<#
.SYNOPSIS
Active Directory - Kullanıcı mail listesi + Mail grupları + Grup üyeleri (nested dahil) Excel raporu.

.DESCRIPTION
- AD'den kullanıcıların mail/UPN/proxyAddresses bilgilerini toplar.
- Mail-enabled grupları (mail veya proxyAddresses veya mailNickname dolu olan grup objeleri) listeler.
- Her grubun üyelerini nested (recursive) şekilde çıkarır.
- Sonuçları tek bir Excel dosyasına 3 sayfa olarak yazar.

.OUTPUTS
.xlsx (ImportExcel modülü ile)

.REQUIREMENTS
- RSAT / ActiveDirectory PowerShell modülü
- ImportExcel PowerShell modülü (Excel kurulu olmak zorunda değil)

.AUTHOR
Ceyhun Yıldız

.DATE
2026-01-20
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -----------------------------
# Ayarlar
# -----------------------------
$Config = [ordered]@{
    RaporKlasoru   = (Join-Path $env:USERPROFILE "Desktop")
    DosyaOnEki     = "AD_Mail_Raporu"
    ExcelAcilsin   = $true

    # Excel görünümü
    TableStyle     = "Medium2"
    FreezeTopRow   = $true
    AutoSize       = $true
    BoldTopRow     = $true

    # Mail grup tespiti (LDAP)
    GroupLdapFilter = "(&(objectCategory=group)(|(mail=*)(proxyAddresses=*)(mailNickname=*)))"
}

$Tarih = Get-Date -Format "yyyy-MM-dd_HH-mm"
$ExcelPath = Join-Path $Config.RaporKlasoru ("{0}_{1}.xlsx" -f $Config.DosyaOnEki, $Tarih)

# -----------------------------
# Yardımcı Fonksiyonlar
# -----------------------------
function Write-Info {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Ok {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-Warn {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

function TS {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return "" }
    return $Text.Trim()
}

function Ensure-Module {
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$InstallIfMissing
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        if ($InstallIfMissing) {
            Write-Warn "$Name modülü bulunamadı. Kurulum deneniyor (CurrentUser)..."
            Install-Module $Name -Scope CurrentUser -Force -AllowClobber
        } else {
            throw "$Name modülü bulunamadı. Lütfen ilgili bileşeni/RSAT'ı kurun."
        }
    }

    Import-Module $Name -ErrorAction Stop
}

function Get-MailEnabledGroups {
    param([Parameter(Mandatory)][string]$LdapFilter)

    $raw = Get-ADGroup -LDAPFilter $LdapFilter -Properties mail, proxyAddresses, mailNickname, groupCategory, groupScope, displayName, managedBy

    return ($raw | Select-Object `
        @{n="GrupAdi";e={TS $_.displayName}},
        @{n="SamAccountName";e={TS $_.sAMAccountName}},
        @{n="GrupMail";e={TS $_.mail}},
        @{n="ProxyAddresses";e={ if ($_.proxyAddresses) { ($_.proxyAddresses | ForEach-Object { TS $_ }) -join "; " } else { "" } }},
        @{n="MailNickname";e={TS $_.mailNickname}},
        @{n="GroupCategory";e={TS $_.groupCategory}},
        @{n="GroupScope";e={TS $_.groupScope}},
        @{n="ManagedByDN";e={TS $_.managedBy}},
        @{n="DN";e={TS $_.DistinguishedName}} |
        Sort-Object GrupAdi)
}

function Get-GroupMembersReport {
    param(
        [Parameter(Mandatory)]$Groups
    )

    foreach ($g in $Groups) {
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
            $uyeTip  = switch ($m.objectClass) {
                "user"  { "Kullanıcı" }
                "group" { "Grup" }
                default { TS $m.objectClass }
            }

            $uyeDisp = TS $m.Name
            $uyeSam  = TS $m.SamAccountName
            $uyeDn   = TS $m.DistinguishedName
            $uyeUpn  = ""
            $uyeMail = ""

            if ($m.objectClass -eq "user") {
                try {
                    $u = Get-ADUser -Identity $m.DistinguishedName -Properties mail, userPrincipalName, displayName, sAMAccountName -ErrorAction Stop
                    $uyeDisp = TS $u.displayName
                    $uyeSam  = TS $u.sAMAccountName
                    $uyeUpn  = TS $u.userPrincipalName
                    $uyeMail = TS $u.mail
                } catch { }
            }
            elseif ($m.objectClass -eq "group") {
                try {
                    $gg = Get-ADGroup -Identity $m.DistinguishedName -Properties mail, displayName, sAMAccountName -ErrorAction Stop
                    $uyeDisp = TS $gg.displayName
                    $uyeSam  = TS $gg.sAMAccountName
                    $uyeMail = TS $gg.mail
                } catch { }
            }

            [pscustomobject]@{
                GrupAdi           = $g.GrupAdi
                GrupMail          = $g.GrupMail
                UyeTipi           = $uyeTip
                UyeDisplayName    = $uyeDisp
                UyeSamAccountName = $uyeSam
                UyeUPN            = $uyeUpn
                UyeMail           = $uyeMail
                UyeDN             = $uyeDn
            }
        }
    }
}

# -----------------------------
# Ön Kontroller
# -----------------------------
Write-Info "Ön kontroller yapılıyor..."

if (-not (Test-Path $Config.RaporKlasoru)) {
    throw "Rapor klasörü bulunamadı: $($Config.RaporKlasoru)"
}

# Modüller
Ensure-Module -Name "ActiveDirectory"
Ensure-Module -Name "ImportExcel" -InstallIfMissing

Write-Ok "Ön kontroller tamam."

# -----------------------------
# Veri Toplama
# -----------------------------
Write-Info "1/3 Kullanıcılar çekiliyor..."
$Kullanicilar = Get-ADUser -Filter * -Properties mail, proxyAddresses, userPrincipalName, displayName, sAMAccountName, enabled |
    Select-Object `
        @{n="DisplayName";e={TS $_.displayName}},
        @{n="SamAccountName";e={TS $_.sAMAccountName}},
        @{n="UPN";e={TS $_.userPrincipalName}},
        @{n="Mail";e={TS $_.mail}},
        @{n="ProxyAddresses";e={ if ($_.proxyAddresses) { ($_.proxyAddresses | ForEach-Object { TS $_ }) -join "; " } else { "" } }},
        @{n="Enabled";e={$_.Enabled}},
        @{n="DN";e={TS $_.DistinguishedName}}

Write-Info "2/3 Mail grupları çekiliyor..."
$Gruplar = Get-MailEnabledGroups -LdapFilter $Config.GroupLdapFilter
Write-Warn ("Bulunan mail-enabled grup sayısı: {0}" -f $Gruplar.Count)

Write-Info "3/3 Grup üyeleri çıkarılıyor (nested dahil)..."
$GrupUyeleri = @(Get-GroupMembersReport -Groups $Gruplar)
Write-Warn ("Toplam üye satırı: {0}" -f ($GrupUyeleri | Measure-Object).Count)

# -----------------------------
# Excel'e Yaz
# -----------------------------
Write-Info "Excel raporu hazırlanıyor..."

if (Test-Path $ExcelPath) {
    Remove-Item $ExcelPath -Force
}

$excelParams = @{
    Path         = $ExcelPath
    AutoSize     = $Config.AutoSize
    FreezeTopRow = $Config.FreezeTopRow
    BoldTopRow   = $Config.BoldTopRow
    TableStyle   = $Config.TableStyle
}

$Kullanicilar | Export-Excel @excelParams -WorksheetName "Kullanicilar"
$Gruplar      | Export-Excel @excelParams -WorksheetName "MailGruplari"
$GrupUyeleri  | Export-Excel @excelParams -WorksheetName "GrupUyeleri"

Write-Ok "Tamamlandı: $ExcelPath"

if ($Config.ExcelAcilsin) {
    Invoke-Item $ExcelPath
}
