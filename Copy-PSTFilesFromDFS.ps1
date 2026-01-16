<#
.SYNOPSIS
Copies PST files from a DFS structure to a local computer.

.DESCRIPTION
Searches a DFS-based file system for PST files and copies them to a specified local directory on the computer.

This script is typically used for collecting Outlook PST files from centralized DFS storage.

.AUTHOR
Ceyhun Yıldız

.DATE
2026-01-16
#>




# Kaynak: DFS yolu veya (tercihen) gerçek share
$source = "\\DOMAIN\DFSRoot\Paylasim"

# Hedef klasör (büyük alan olmalı)
$dest   = "D:\PST_Collect"

# Log
$logDir = "D:\PST_Collect_Logs"
New-Item -ItemType Directory -Force -Path $dest, $logDir | Out-Null
$log = Join-Path $logDir ("pst_copy_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

# Kopyalama: sadece *.pst
# /S: alt klasörler
# /MT: çoklu thread (16/32/64 deneyebilirsin)
# /R:0 /W:0: takılan dosyada beklemesin
# /XO: hedefte daha yeni varsa geç
# /XJ: junction'lara girme (loop riskini azaltır)
# /NP: yüzde progress basma (log şişmesin)
robocopy $source $dest *.pst /S /MT:32 /R:0 /W:0 /XO /XJ /COPY:DAT /DCOPY:DAT /NP /TEE /LOG:$log

Write-Host "Bitti. Log: $log"


# “Kopyalamadan önce listeleyeyim” (hızlı ön kontrol)
# Önce kaç tane var görmek istersen:
# $source = "\\DOMAIN\DFSRoot\Paylasim"
# $dest   = "D:\PST_Collect"
# robocopy $source $dest *.pst /S /L /XJ /R:0 /W:0
