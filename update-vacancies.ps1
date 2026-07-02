# Vermaak PA – Auto Vacancy Updater
# Reads the latest vacancy report .xlsx and injects data into pa-command-centre.html
# Scheduled: Wednesday & Friday at 04:00

$reportFolder = "C:\Users\02jus\OneDrive - Vermaak Properties\Desktop\Vacancy Reports 2026"
$htmlFile     = "C:\Users\02jus\OneDrive - Vermaak Properties\Desktop\Apps & Websites\pa-command-centre.html"
$logFile      = "C:\Users\02jus\OneDrive - Vermaak Properties\Desktop\Apps & Websites\update-vacancies.log"

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content $logFile "[$ts] $msg"
}

Log "--- Run started ---"

# Find the most recently modified vacancy report
$latest = Get-ChildItem $reportFolder -Filter "Vermaak Vacancy Report*.xlsx" -ErrorAction SilentlyContinue |
          Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (!$latest) { Log "ERROR: No vacancy report found in $reportFolder"; exit 1 }
Log "Reading: $($latest.Name)"

# Open with Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb    = $excel.Workbooks.Open($latest.FullName)
$sheet = $wb.Sheets.Item(1)
$rows  = $sheet.UsedRange.Rows.Count

$locMap   = @{ 'Tyger Valley'='Tygervalley'; 'Century City'='Century City'; 'Bellville'='Bellville'; 'Plattekloof'='Plattekloof' }
$monMap   = @{ 'Jan'='01';'Feb'='02';'Mar'='03';'Apr'='04';'May'='05';'Jun'='06';'Jul'='07';'Aug'='08';'Sep'='09';'Oct'='10';'Nov'='11';'Dec'='12' }

# Find header row (column A = "Company")
$headerRow = -1
for ($r = 1; $r -le $rows; $r++) {
    if ($sheet.Cells($r, 1).Text.Trim().ToLower() -eq 'company') { $headerRow = $r; break }
}
if ($headerRow -lt 0) { Log "ERROR: Header row not found"; $wb.Close($false); $excel.Quit(); exit 1 }

$vacLines = @()
for ($r = $headerRow + 1; $r -le $rows; $r++) {
    $company  = $sheet.Cells($r, 1).Text.Trim()
    $building = $sheet.Cells($r, 2).Text.Trim()
    if (!$building -or $company.ToLower().StartsWith('report generated')) { break }

    $location    = $sheet.Cells($r,  3).Text.Trim()
    $sizeRaw     = $sheet.Cells($r,  4).Text.Trim().Replace(',', '.')
    $sizeInt     = ($sizeRaw -split '\.')[0]
    $parkBase    = $sheet.Cells($r,  5).Text.Trim()
    $parkCover   = $sheet.Cells($r,  6).Text.Trim()
    $parkOpen    = $sheet.Cells($r,  7).Text.Trim()
    $ac          = $sheet.Cells($r,  8).Text.Trim()
    $security    = $sheet.Cells($r,  9).Text.Trim()
    $rentalRaw   = $sheet.Cells($r, 10).Text.Trim().Replace(',', '.')
    $schedDate   = $sheet.Cells($r, 11).Text.Trim()

    $area   = if ($locMap.ContainsKey($location)) { $locMap[$location] } else { 'Bellville' }
    $rental = if ($rentalRaw) { "R$rentalRaw/m``u00b2" } else { '' }

    $vacDate = ''
    $dp = $schedDate -split ' '
    if ($dp.Count -eq 2 -and $monMap.ContainsKey($dp[0])) {
        $vacDate = "$($dp[1])-$($monMap[$dp[0]])-01"
    }

    $noteParts = @()
    if ($company)   { $noteParts += "Owner: $company" }
    if ($parkBase)  { $noteParts += "Basement: $parkBase" }
    if ($parkCover) { $noteParts += "Covered: $parkCover" }
    if ($parkOpen)  { $noteParts += "Open: $parkOpen" }
    if ($ac)        { $noteParts += "AC: $ac" }
    if ($security)  { $noteParts += "Security: $security" }
    $notes = $noteParts -join ' | '

    # Escape single quotes for JS
    $b = $building -replace "'", "\'"
    $n = $notes    -replace "'", "\'"
    $ren = $rental -replace "'", "\'"

    $vacLines += "  { building: '$b', tenant: '', area: '$area', size: '$sizeInt', rental: '$ren', vacancyDate: '$vacDate', status: 'Vacant', notes: '$n' },"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

if ($vacLines.Count -eq 0) { Log "ERROR: No vacancy rows parsed"; exit 1 }

$reportDate = Get-Date -Format "yyyy-MM-dd"
$newBlock   = "const VACANCY_REPORT_DATE = '$reportDate';`r`nconst VACANCY_DEFAULT_DATA = [`r`n" + ($vacLines -join "`r`n") + "`r`n];"

$html    = [System.IO.File]::ReadAllText($htmlFile, [System.Text.Encoding]::UTF8)
$pattern = "(?s)const VACANCY_REPORT_DATE = '[^']*';\s*const VACANCY_DEFAULT_DATA = \[.*?\];"
$html    = [System.Text.RegularExpressions.Regex]::Replace($html, $pattern, $newBlock)

[System.IO.File]::WriteAllText($htmlFile, $html, [System.Text.Encoding]::UTF8)
Log "SUCCESS: Wrote $($vacLines.Count) vacancies from $($latest.Name)"
