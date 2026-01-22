<# 
.SYNOPSIS
Active Directory - Exports all users (including nested members) of one or more selected groups (comma-separated) to Excel with Turkish column names and opens the file automatically.

.REQUIREMENTS
- RSAT / ActiveDirectory PowerShell module
- ImportExcel PowerShell module (Microsoft Excel is not required)

.AUTHOR
Ceyhun Yıldız
.DATE
2026-01-21
#>

[CmdletBinding()]
param()

# =========================
# AYARLAR (Sadece burayı değiştirmen yeterli)
# =========================
$OutputDir  = "C:\AD_PS_Ops\Reports"
$FileName   = "AD_MultiGroup_Members_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$OutputPath = Join-Path $OutputDir $FileName
# =========================

# --- Çıktı dizini kontrol ---
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

# --- Ön kontroller ---
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Host "ActiveDirectory modülü bulunamadı. RSAT yüklü mü kontrol et." -ForegroundColor Red
    throw
}

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel modülü yok. Kuruluyor (CurrentUser)..." -ForegroundColor Yellow
    try {
        Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
    } catch {
        Write-Host "ImportExcel kurulamadı. PowerShell'i Yönetici olarak açıp tekrar deneyebilirsin." -ForegroundColor Red
        throw
    }
}
Import-Module ImportExcel -ErrorAction Stop

# --- Grupları sor ---
$groupInputRaw = Read-Host "Rapor almak istediğin AD gruplarını yaz (virgülle ayır): (isim/CN/SamAccountName/DN)"
if ([string]::IsNullOrWhiteSpace($groupInputRaw)) {
    Write-Host "Grup adı boş olamaz." -ForegroundColor Red
    exit 1
}

$groupInputs = $groupInputRaw.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }

if (-not $groupInputs -or $groupInputs.Count -eq 0) {
    Write-Host "Geçerli grup girişi bulunamadı." -ForegroundColor Red
    exit 1
}

# Kullanıcıları tekilleştirmek ve grup bilgisini tutmak için
$userMap = @{}

# --- Her grup için üyeleri çek ---
foreach ($groupInput in $groupInputs) {

    # Grubu çöz
    $group = $null
    try {
        $group = Get-ADGroup -Identity $groupInput -ErrorAction Stop
    } catch {
        $group = Get-ADGroup -Filter "Name -eq '$groupInput' -or SamAccountName -eq '$groupInput'" -ErrorAction SilentlyContinue | Select-Object -First 1
    }

    if (-not $group) {
        Write-Host "Grup bulunamadı, atlanıyor: $groupInput" -ForegroundColor Yellow
        continue
    }

    Write-Host "İşleniyor -> Grup: $($group.Name)" -ForegroundColor Cyan

    # Üyeleri çek (nested dahil)
    try {
        $members = Get-ADGroupMember -Identity $group.DistinguishedName -Recursive -ErrorAction Stop
    } catch {
        Write-Host "Grup üyeleri alınamadı, atlanıyor: $($group.Name)" -ForegroundColor Yellow
        continue
    }

    $userMembers = $members | Where-Object { $_.objectClass -eq 'user' }

    foreach ($u in $userMembers) {
        try {
            $adUser = Get-ADUser -Identity $u.DistinguishedName -Properties `
                DisplayName, SamAccountName, UserPrincipalName, Mail, TelephoneNumber, MobilePhone, `
                Title, Department, Company, Manager, Enabled, LastLogonDate, `
                whenCreated, whenChanged, DistinguishedName

            # Manager çöz
            $managerName = $null
            if ($adUser.Manager) {
                try {
                    $managerName = (Get-ADUser -Identity $adUser.Manager -Properties DisplayName).DisplayName
                } catch {
                    $managerName = $adUser.Manager
                }
            }

            $key = $adUser.SamAccountName

            if ($userMap.ContainsKey($key)) {
                $existingGroups = $userMap[$key]."Üye Olduğu Gruplar" -split '\s*;\s*'
                if ($existingGroups -notcontains $group.Name) {
                    $userMap[$key]."Üye Olduğu Gruplar" = (($existingGroups + $group.Name) | Sort-Object) -join "; "
                }
            }
            else {
                $userMap[$key] = [PSCustomObject]@{
                    "Ad Soyad"               = $adUser.DisplayName
                    "Kullanıcı Adı"          = $adUser.SamAccountName
                    "UPN"                    = $adUser.UserPrincipalName
                    "E-Posta"                = $adUser.Mail
                    "Telefon"                = $adUser.TelephoneNumber
                    "Cep Telefonu"           = $adUser.MobilePhone
                    "Unvan"                  = $adUser.Title
                    "Departman"              = $adUser.Department
                    "Şirket"                 = $adUser.Company
                    "Yönetici"               = $managerName
                    "Hesap Aktif"            = $adUser.Enabled
                    "Son Oturum Açma"        = $adUser.LastLogonDate
                    "Oluşturulma Tarihi"     = $adUser.whenCreated
                    "Son Değişiklik Tarihi"  = $adUser.whenChanged
                    "DN"                     = $adUser.DistinguishedName
                    "Üye Olduğu Gruplar"     = $group.Name
                }
            }
        } catch {
            Write-Host "Kullanıcı okunamadı: $($u.Name)" -ForegroundColor Yellow
        }
    }
}

# --- Sonuçlar ---
$usersDetailed = $userMap.Values
if (-not $usersDetailed) {
    Write-Host "Raporlanacak kullanıcı bulunamadı." -ForegroundColor Yellow
    exit 0
}

# --- Excel'e yaz ---
$usersDetailed |
    Sort-Object "Ad Soyad" |
    Export-Excel -Path $OutputPath `
        -WorksheetName "Grup Üyeleri" `
        -TableName "GrupUyeleri" `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -ClearSheet

Write-Host "Excel oluşturuldu: $OutputPath" -ForegroundColor Green

# --- Otomatik aç ---
Invoke-Item -Path $OutputPath
