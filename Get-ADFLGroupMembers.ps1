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




Import-Module ActiveDirectory -ErrorAction Stop
$ErrorActionPreference = "Stop"

# Paths
$ReportsDir = "C:\AD_PS_Ops\Reports"
if (-not (Test-Path $ReportsDir)) {
    New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null
}

$Suffix    = "FL"
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvPath   = Join-Path $ReportsDir "FL_Groups_$TimeStamp.csv"
$HtmlPath  = Join-Path $ReportsDir "FL_Groups_$TimeStamp.html"

# Helper
function Safe($v) {
    if ($null -eq $v) { return "" }
    return [string]$v
}

# Recursive group expansion
function Get-NestedMembers {
    param(
        [string]$GroupDN,
        [string]$TopGroupName,
        [string]$Path,
        [hashtable]$Visited
    )

    if (-not $Visited) { $Visited = @{} }
    if ($Visited.ContainsKey($GroupDN)) { return @() }
    $Visited[$GroupDN] = $true

    $rows = @()

    try {
        $members = Get-ADGroupMember -Identity $GroupDN
    } catch {
        return @()
    }

    foreach ($m in $members) {

        if ($m.objectClass -eq "group") {
            $newPath = "$Path > $($m.Name)"
            $rows += Get-NestedMembers -GroupDN $m.DistinguishedName -TopGroupName $TopGroupName -Path $newPath -Visited $Visited
        }
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

# ================= MAIN =================

Write-Host "Aranıyor: Adı veya SamAccountName'i '$Suffix' ile biten gruplar..." -ForegroundColor Cyan

$groups = Get-ADGroup -Filter "Name -like '*$Suffix' -or SamAccountName -like '*$Suffix'" `
            -Properties Name,DistinguishedName

$groups = @($groups)   # 🔑 GARANTİ ARRAY

Write-Host ("Bulunan grup sayısı: {0}" -f $groups.Length) -ForegroundColor Yellow

$allRows = @()

foreach ($g in $groups) {
    Write-Host "İşleniyor: $($g.Name)" -ForegroundColor DarkGray
    $visited = @{}
    $allRows += Get-NestedMembers `
        -GroupDN $g.DistinguishedName `
        -TopGroupName $g.Name `
        -Path $g.Name `
        -Visited $visited
}

# CSV
$allRows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

# HTML (basit, sağlam)
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
<td>$($r.GrupAdi)</td>
<td>$($r.Uye)</td>
<td>$($r.Hesap)</td>
<td>$($r.Mail)</td>
<td>$($r.Enabled)</td>
<td>$($r.Path)</td>
</tr>"
}

$html += "</table></body></html>"

Set-Content -Path $HtmlPath -Value $html -Encoding UTF8

Write-Host ""
Write-Host "TAMAMLANDI ✅" -ForegroundColor Green
Write-Host "CSV  : $CsvPath"
Write-Host "HTML : $HtmlPath"
Write-Host ""

# CSV dosyasını Excel ile otomatik aç
if (Test-Path $CsvPath) {
    Invoke-Item -Path $CsvPath
}

