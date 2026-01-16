<#
.SYNOPSIS
Exports Active Directory users created in the last 7 days to Excel.

.DESCRIPTION
Retrieves Active Directory users created within the last 7 days and exports
the following attributes to an Excel file:
- SamAccountName
- DisplayName
- Mail
- Description
- POBox

The generated Excel file is automatically opened after export.

.AUTHOR
Ceyhun Yıldız

.DATE
2026-01-16
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Ön kontroller ---
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "ActiveDirectory modülü bulunamadı. RSAT/AD module kurulu olmalı."
}
Import-Module ActiveDirectory -ErrorAction Stop

# --- Filtre zamanı ---
$since = (Get-Date).AddDays(-7)

# --- Veriyi çek ---
$users = Get-ADUser -Filter { whenCreated -ge $since } -Properties DisplayName, Mail, Description, POBox, whenCreated |
    Select-Object `
        @{n='Kullanıcı Adı'; e={$_.SamAccountName}},
        @{n='Görünen Ad';   e={$_.DisplayName}},
        @{n='Mail Adresi';  e={$_.Mail}},
        @{n='Açıklama';     e={$_.Description}},
        @{n='Posta Kutusu'; e={$_.POBox}},
        @{n='Oluşturulma';  e={$_.whenCreated}} |
    Sort-Object 'Oluşturulma' -Descending   # ✅ DÜZELTİLDİ

# --- Çıktı dosyası ---
$outDir = "C:\Temp"
if (-not (Test-Path $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }

$outFile = Join-Path $outDir ("AD_Son7Gun_Kullanicilar_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

# --- Excel'e yaz (COM) ---
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    throw "Excel COM objesi oluşturulamadı. Bu makinede Microsoft Excel yüklü olmayabilir."
}

$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)
$sheet.Name = "Son7Gun"

try {
    # Başlıklar
    $headers = @("Kullanıcı Adı","Görünen Ad","Mail Adresi","Açıklama","Posta Kutusu","Oluşturulma")
    for ($c=0; $c -lt $headers.Count; $c++) {
        $sheet.Cells.Item(1, $c+1).Value2 = $headers[$c]
    }

    # Satırlar
    $row = 2
    foreach ($u in $users) {
        $sheet.Cells.Item($row,1).Value2 = $u.'Kullanıcı Adı'
        $sheet.Cells.Item($row,2).Value2 = $u.'Görünen Ad'
        $sheet.Cells.Item($row,3).Value2 = $u.'Mail Adresi'
        $sheet.Cells.Item($row,4).Value2 = $u.'Açıklama'
        $sheet.Cells.Item($row,5).Value2 = $u.'Posta Kutusu'
        $sheet.Cells.Item($row,6).Value2 = ($u.'Oluşturulma').ToString("yyyy-MM-dd HH:mm:ss")
        $row++
    }

    # Basit biçimlendirme
    $used = $sheet.UsedRange
    $used.EntireColumn.AutoFit() | Out-Null
    $sheet.Range("A1:F1").Font.Bold = $true
    $sheet.Range("A1:F1").AutoFilter() | Out-Null
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    # Kaydet/Kapat
    $workbook.SaveAs($outFile)
    $workbook.Close($true)
    $excel.Quit()
}
finally {
    # ✅ Hata olsa bile COM temizliği garanti
    if ($sheet)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)    | Out-Null }
    if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)    | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Excel'i aç
if (Test-Path $outFile) {
    Start-Process $outFile
    Write-Host "Excel oluşturuldu ve açıldı: $outFile" -ForegroundColor Green
} else {
    Write-Host "Dosya oluşturulamadı: $outFile" -ForegroundColor Red
}
