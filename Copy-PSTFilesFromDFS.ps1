<#
.SYNOPSIS
Copies PST files from a DFS structure to a local computer.

.DESCRIPTION
Searches a DFS-based file system for PST files and copies them
to a specified local directory using robocopy.

This script is typically used for collecting Outlook PST files
from centralized DFS storage.

.AUTHOR
Ceyhun Yıldız

.DATE
2026-01-16
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =========================
# AYARLAR (Ortamına göre düzenle)
# =========================

# DFS veya file share yolu (ÖRNEK)
$SourcePath = "\\YOURDOMAIN\DFS\Share"

# PST dosyalarının kopyalanacağı klasör
$DestinationPath = "C:\AD_PS_Ops\PST"

# Log klasörü
$LogDir = Join-Path $DestinationPath "Logs"

# =========================
# KOD
# =========================

# Klasörleri oluştur
New-Item -ItemType Directory -Force -Path $DestinationPath, $LogDir | Out-Null

# Log dosyası
$LogFile = Join-Path $LogDir ("pst_copy_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

# Robocopy parametreleri
# /S    : Alt klasörler
# /MT   : Çoklu thread
# /R:0  : Retry yok
# /W:0  : Bekleme yok
# /XO   : Hedefte yeni dosya varsa geç
# /XJ   : Junction'lara girme
# /NP   : Progress gösterme
robocopy `
    $SourcePath `
    $DestinationPath `
    *.pst `
    /S /MT:32 /R:0 /W:0 /XO /XJ /COPY:DAT /DCOPY:DAT /NP /TEE /LOG:$LogFile

Write-Host "TAMAMLANDI ✅" -ForegroundColor Green
Write-Host "Log dosyası: $LogFile"

# =========================
# ÖN KONTROL (isteğe bağlı)
# =========================
# Sadece listelemek için:
# robocopy $SourcePath $DestinationPath *.pst /S /L /XJ /R:0 /W:0
