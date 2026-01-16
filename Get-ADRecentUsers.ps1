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

# Yazım hatalarını ve beklenmedik durumları daha erken yakalamak için
Set-StrictMode -Version Latest

# Hata olduğunda devam etmesin, direkt dursun
$ErrorActionPreference = "Stop"

# =========================
# AYARLAR (Sadece burayı değiştirmen yeterli)
# =========================

# Kaç gün geriye giderek kullanıcıları çeksin?
$DaysBack = 7

# Çıktıların kaydedileceği klasör (repo standardı)
$OutputDir = "C:\AD_PS_Ops\Reports"

# Export bittiğinde Excel otomatik açılsın mı?
$OpenAfterExport = $true

# =========================
# ÖN KONTROLLER
# =========================

# ActiveDirectory modülü yüklü mü? (RSAT gerekli)
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "ActiveDirectory modülü bulunamadı. RSAT/AD module kurulu olmalı."
}

# AD modülünü yükle
Import-Module ActiveDirectory -ErrorAction Stop

# Output klasörü yoksa oluştur
New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null

# =========================
# VERİ ÇEKME
# =========================

# Filtre zamanı: bugün - DaysBack
$since = (Get-Date).AddDays(-$DaysBack)

# Son X günde oluşturulan kullanıcıları çek
$users = Get-ADUser -Filter { whenCreated -ge $since } -Properties DisplayName, Mail, Description, POBox, whenCreated |
    Select-Object `
        @{n='Kullanıcı Adı'; e={$_.SamAccountName}},
        @{n='Görünen Ad';   e={$_.DisplayName}},
        @{n='Mail Adresi';  e={$_.Mail}},
        @{n='Açıklama';     e={$_.Description}},
        @{n='Posta Kutusu'; e={$_.POBox}},
        @{n='Oluşturulma';  e={$_.whenCreated}} |
    Sort-Object 'Oluşturulma' -Descending

# ✅ KRİTİK KORUMA:
# Hiç kullanıcı yoksa Excel COM’a girip gereksiz işlem yapma
if (-not $users -or $users.Count -eq 0) {
    Write-Host "Son $DaysBack gün içinde oluşturulan kullanıcı bulunamadı." -ForegroundColor Yellow
    return
}

# =========================
# ÇIKTI DOSYASI
# =========================

# Zaman damgası ile benzersiz dosya adı
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Excel dosyası yolu
$outFile = Join-Path $OutputDir ("AD_Son{0}Gun_Kullanicilar_{1}.xlsx" -f $DaysBack, $TimeStamp)

# =========================
# EXCEL'E YAZ (COM)
# =========================

# Excel COM objesi oluştur
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    throw "Excel COM objesi oluşturulamadı. Bu makinede Microsoft Excel yüklü olmayabilir."
}

# Excel arka planda çalışsın
$excel.Visible = $false

# Yeni workbook ve sheet oluştur
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)
$sheet.Name = "Son$DaysBack" + "Gun"

try {
    # Başlıklar (kolon isimleri)
    $headers = @("Kullanıcı Adı","Görünen Ad","Mail Adresi","Açıklama","Posta Kutusu","Oluşturulma")

    # Başlık satırını yaz
    for ($c=0; $c -lt $headers.Count; $c++) {
        $sheet.Cells.Item(1, $c+1).Value2 = $headers[$c]
    }

    # Verileri yaz (2. satırdan itibaren)
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

    # Biçimlendirme: kolonları sığdır, başlığı bold yap, filtre ekle, başlığı dondur
    $used = $sheet.UsedRange
    $used.EntireColumn.AutoFit() | Out-Null
    $sheet.Range("A1:F1").Font.Bold = $true
    $sheet.Range("A1:F1").AutoFilter() | Out-Null
    $excel.ActiveWindow.SplitRow = 1
    $excel.ActiveWindow.FreezePanes = $true

    # Kaydet ve kapat
    $workbook.SaveAs($outFile)
    $workbook.Close($true)
    $excel.Quit()
}
finally {
    # COM temizliği: Excel process arkada kalmasın
    if ($sheet)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)    | Out-Null }
    if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
    if ($excel)    { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)    | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# =========================
# DOSYAYI AÇ / BİLGİ
# =========================

# Excel dosyasını otomatik aç
if ($OpenAfterExport -and (Test-Path $outFile)) {
    Start-Process $outFile
}

# Kullanıcıya bilgi ver
if (Test-Path $outFile) {
    Write-Host "Excel oluşturuldu ve açıldı: $outFile" -ForegroundColor Green
} else {
    Write-Host "Dosya oluşturulamadı: $outFile" -ForegroundColor Red
}
