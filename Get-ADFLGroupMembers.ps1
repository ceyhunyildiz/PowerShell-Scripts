<#
.SYNOPSIS
Retrieves members of Active Directory groups ending with "FL".

.DESCRIPTION
Identifies Active Directory groups whose names end with "FL" and retrieves
their members, including nested group memberships.

This script is commonly used to analyze role-based or functional
group structures in Active Directory environments.

.AUTHOR
Ceyhun Yıldız

.DATE
2026-01-16
#>

# Yazım hataları / tanımsız değişken gibi durumlarda daha erken hata yakalamak için
Set-StrictMode -Version Latest

# Hata olduğunda scriptin devam etmesini istemiyoruz; direkt dursun
$ErrorActionPreference = "Stop"

# =========================
# AYARLAR (Sadece burayı değiştirmen yeterli)
# =========================

# Raporların kaydedileceği ana klasör
$OutputDir = "C:\AD_PS_Ops\Reports"

# Sonu hangi ifadeyle biten gruplar aranacak?
$Suffix = "FL"

# CSV tamamlanınca otomatik açılsın mı?
$OpenCsvAfterExport = $true

# (Opsiyonel) Belirli bir OU içindeki gruplarda aramak istersen:
# Örn: "OU=Groups,DC=domain,DC=local"
$SearchBase = ""

# =========================
# ÖN KONTROLLER
# =========================

# ActiveDirectory modülü yüklü mü kontrol et (RSAT kurulu olmalı)
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "ActiveDirectory modülü bulunamadı. RSAT (Active Directory module) kurulu olmalı."
}

# Modülü yükle (hata olursa Stop)
Import-Module ActiveDirectory -ErrorAction Stop

# Output klasörü yoksa oluştur
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

# Zaman damgası (dosya adına eklenecek)
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Çıktı dosyaları
$CsvPath  = Join-Path $OutputDir "FL_Groups_$TimeStamp.csv"
$HtmlPath = Join-Path $OutputDir "FL_Groups_$TimeStamp.html"

# =========================
# YARDIMCI FONKSİYONLAR
# =========================

# Null değerleri boş string’e çevir (CSV/HTML’de daha temiz görünür)
function Safe {
    param([object]$v)
    if ($null -eq $v) { return "" }
    return [string]$v
}

# HTML içine basarken özel karakterleri encode et (HTML kırılmasını engeller)
function HtmlEncode {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($s)
}

# Nested (iç içe) grup üyeliklerini recursive çözen fonksiyon
function Get-NestedMembers {
    param(
        [string]$GroupDN,         # İşlenecek grubun DN'i
        [string]$TopGroupName,    # En üst (raporda görünen) grup adı
        [string]$Path,            # Üyeliğin yolu (hangi gruplardan geçti)
        [hashtable]$Visited       # Loop engelleme (aynı gruba tekrar girme)
    )

    # Ziyaret tablosu yoksa oluştur
    if (-not $Visited) { $Visited = @{} }

    # Aynı gruba tekrar girmeyi engelle (loop riskini azaltır)
    if ($Visited.ContainsKey($GroupDN)) { return @() }
    $Visited[$GroupDN] = $true

    $rows = @()

    # Grup üyelerini çek (hata olursa boş dön)
    try {
        $members = Get-ADGroupMember -Identity $GroupDN
    } catch {
        return @()
    }

    foreach ($m in $members) {

        # Üye başka bir grupsa, onun içine de gir (nested)
        if ($m.objectClass -eq "group") {
            $newPath = "$Path > $($m.Name)"
            $rows += Get-NestedMembers `
                -GroupDN $m.DistinguishedName `
                -TopGroupName $TopGroupName `
                -Path $newPath `
                -Visited $Visited
        }
        # Üye kullanıcıysa detaylarını çek ve satır üret
        elseif ($m.objectClass -eq "user") {
            $u = Get-ADUser -Identity $m.DistinguishedName -Properties DisplayName,Mail,Enabled -ErrorAction SilentlyContinue

            $rows += [pscustomobject]@{
                GrupAdi = $TopGroupName
                Uye     = if ($u.DisplayName) { $u.DisplayName } else { $m.Name }
                Hesap   = Safe $u.SamAccountName
                Mail    = Safe $u.Mail
                Enabled = Safe $u.Enabled
                Path    = $Path
            }
        }
    }

    return $rows
}

# =========================
# MAIN
# =========================

Write-Host "Aranıyor: Adı veya SamAccountName'i '$Suffix' ile biten gruplar..." -ForegroundColor Cyan

# Grup araması parametreleri
$groupParams = @{
    Filter     = "Name -like '*$Suffix' -or SamAccountName -like '*$Suffix'"
    Properties = @("Name","DistinguishedName")
}

# SearchBase verilmişse grupları o OU altında arar (opsiyonel)
if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
    $groupParams.SearchBase = $SearchBase
    Write-Host "Bilgi: SearchBase kullanılıyor -> $SearchBase" -ForegroundColor Yellow
}

# Grupları çek
$groups = Get-ADGroup @groupParams

# Null veya tek öğe gelme ihtimaline karşı array garantisi
$groups = @($groups)

Write-Host ("Bulunan grup sayısı: {0}" -f $groups.Length) -ForegroundColor Yellow

# Hiç grup yoksa boş rapor üretip çık
if ($groups.Length -eq 0) {
    # Boş CSV
    @() | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

    # Boş HTML
    $emptyHtml = @"
<html>
<head><meta charset='utf-8'><title>FL Groups Report</title></head>
<body style="font-family:Segoe UI;">
<h2>FL Grupları – Nested Üyeler</h2>
<p>Hiç grup bulunamadı. (Suffix: $Suffix)</p>
</body></html>
"@
    Set-Content -Path $HtmlPath -Value $emptyHtml -Encoding UTF8

    Write-Host "Hiç grup bulunamadı. Boş raporlar oluşturuldu." -ForegroundColor Yellow
    Write-Host "CSV  : $CsvPath"
    Write-Host "HTML : $HtmlPath"
    return
}

# Tüm grupların üyelerini toplamak için
$allRows = @()

foreach ($g in $groups) {
    Write-Host "İşleniyor: $($g.Name)" -ForegroundColor DarkGray

    # Her grup için ayrı visited tablosu (loop önleme)
    $visited = @{}

    # Nested üyeleri çek
    $allRows += Get-NestedMembers `
        -GroupDN $g.DistinguishedName `
        -TopGroupName $g.Name `
        -Path $g.Name `
        -Visited $visited
}

# CSV export (Excel uyumlu)
$allRows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

# =========================
# HTML RAPOR
# =========================

$html = @"
<html>
<head>
<meta charset='utf-8'>
<title>FL Groups Report</title>
<style>
body{font-family:Segoe UI;background:#0b1220;color:#eaefff}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #2a3a5e;padding:8px}
th{background:#142040}
tr:nth-child(even){background:#101a30}
</style>
</head>
<body>
<h2>FL Grupları – Nested Üyeler</h2>
<table>
<tr>
<th>Grup</th>
<th>Üye</th>
<th>Hesap</th>
<th>Mail</th>
<th>Enabled</th>
<th>Path</th>
</tr>
"@

foreach ($r in $allRows) {
    $html += "<tr>
<td>$(HtmlEncode (Safe $r.GrupAdi))</td>
<td>$(HtmlEncode (Safe $r.Uye))</td>
<td>$(HtmlEncode (Safe $r.Hesap))</td>
<td>$(HtmlEncode (Safe $r.Mail))</td>
<td>$(HtmlEncode (Safe $r.Enabled))</td>
<td>$(HtmlEncode (Safe $r.Path))</td>
</tr>"
}

$html += "</table></body></html>"

# HTML dosyasını yaz
Set-Content -Path $HtmlPath -Value $html -Encoding UTF8

# =========================
# SONUÇ / AÇMA
# =========================

Write-Host ""
Write-Host "TAMAMLANDI ✅" -ForegroundColor Green
Write-Host "CSV  : $CsvPath"
Write-Host "HTML : $HtmlPath"
Write-Host ("Toplam satır: {0}" -f @($allRows).Count) -ForegroundColor Yellow
Write-Host ""

# CSV dosyasını otomatik aç (Excel varsayılan uygulama olarak açar)
if ($OpenCsvAfterExport -and (Test-Path $CsvPath)) {
    Invoke-Item -Path $CsvPath
}
