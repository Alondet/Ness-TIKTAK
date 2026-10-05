#requires -version 5.1
<#
Attendance -> Timesheet converter
Requires Microsoft Excel installed (uses Excel COM automation).

Expected source: EmpMonthTotals.xlsx
Expected template: Timesheet Report Template.xls

Template layout used:
  Row 7     = regular work (M01192 / 1001)
  Rows 11-14 = extra regular-work intervals (created automatically from row 7)
  Rows 8-10 = vacation / sick / admin - left untouched for manual editing
#>

[CmdletBinding()]
param(
    [string]$AttendanceFile,
    [string]$TemplateFile,
    [string]$OutputFile
)

$ErrorActionPreference = 'Stop'

# ---------- AlonDet Console UI ----------
$Host.UI.RawUI.BackgroundColor = 'Black'
$Host.UI.RawUI.ForegroundColor = 'White'
try { $Host.UI.RawUI.WindowTitle = 'AlonDet - Attendance -> Timesheet' } catch {}

# Exact AlonDet logo from IT_Toolkit_Menu(4).ps1 (TAAG Breach Blue encoding).
$script:Logo = @(
    ' DDD  DD              DDDD        DD   '
    'BBUBB BB              BBUBB       BB   '
    'HH HH HH              HH HH       HH   '
    'HH HH HH              HH HH       HH   '
    'MM MM MM   DDD  DDDD  MM MM  DDDD MMMM '
    'MMMMM MM  MM MM MM MM MM MM MM MM MM   '
    'LL LL LL  LL LL LL LL LL LL LLDLL LL   '
    'LL LL LL  LL LL LL LL LL LL LL DD LL   '
    'BB BB BB  BB BB BB BB BBDBB BB BB BB B '
    'UU UU  UU  UUU  UU UU UUUU   UUUU  UUU '
)

function Write-AlonDetBannerRow {
    param([string]$Encoded, [int]$Row, [int]$Width = 39)
    foreach ($run in [regex]::Matches($Encoded, '[HML]+|[DBU]+| +')) {
        $value = $run.Value
        $foreground = 'Magenta'
        $background = 'Black'
        if ($value[0] -in @('H','M','L')) { $background = 'DarkMagenta' }
        elseif ($value[0] -ne ' ' -and $Row -ge 6) { $foreground = 'DarkMagenta' }
        $value = $value.Replace([char]'D', [char]0x2584).Replace([char]'U', [char]0x2580).Replace([char]'B', [char]0x2588)
        $value = $value.Replace([char]'H', [char]0x2593).Replace([char]'M', [char]0x2592).Replace([char]'L', [char]0x2591)
        Write-Host $value -ForegroundColor $foreground -BackgroundColor $background -NoNewline
    }
    if ($Width -gt $Encoded.Length) { Write-Host (' ' * ($Width - $Encoded.Length)) -BackgroundColor Black -NoNewline }
    Write-Host ''
}

function Show-AlonHeader {
    Clear-Host
    Write-Host ('=' * 72) -ForegroundColor DarkMagenta
    for ($i=0; $i -lt $script:Logo.Count; $i++) { Write-AlonDetBannerRow $script:Logo[$i] $i 39 }
    Write-Host ''
    Write-Host '  ATTENDANCE -> TIMESHEET' -ForegroundColor Magenta
    Write-Host '  CONSOLE EDITION' -ForegroundColor DarkGray -NoNewline
    Write-Host ('v.6'.PadLeft(53)) -ForegroundColor Magenta
    Write-Host ('=' * 72) -ForegroundColor DarkMagenta
    Write-Host ''
    $script:ProgressInitialized = $false
}

# Custom in-console progress. Unlike Write-Progress, it stays below the AlonDet logo.
function Show-RealProgress([int]$Percent,[string]$Status) {
    $Percent = [Math]::Max(0,[Math]::Min(100,$Percent))
    $width = 42
    $filled = [int][Math]::Floor($width * ($Percent / 100.0))
    $empty = $width - $filled
    $bar = ('#' * $filled) + ('-' * $empty)

    if (-not $script:ProgressInitialized) {
        Write-Host '  PROCESSING' -ForegroundColor White
        Write-Host ''
        Write-Host ''
        Write-Host ''
        $script:ProgressTop = $Host.UI.RawUI.CursorPosition.Y - 3
        $script:ProgressInitialized = $true
    }

    try {
        $pos = $Host.UI.RawUI.CursorPosition
        $pos.X = 0; $pos.Y = $script:ProgressTop
        $Host.UI.RawUI.CursorPosition = $pos
        Write-Host '  [' -ForegroundColor DarkGray -NoNewline
        Write-Host ('#' * $filled) -ForegroundColor Cyan -NoNewline
        Write-Host ('-' * $empty) -ForegroundColor DarkGray -NoNewline
        Write-Host ('] {0,3}%' -f $Percent) -ForegroundColor White
        Write-Host ('  Current: {0}' -f $Status).PadRight(70) -ForegroundColor Gray
        Write-Host ('  ' + ('-' * 68)) -ForegroundColor DarkMagenta
    }
    catch {
        Write-Host ("  [$bar] $Percent%  $Status") -ForegroundColor Cyan
    }
}

Show-AlonHeader
Write-Host '  Ready. Select the attendance file and the timesheet template.' -ForegroundColor White
Write-Host ''

# ---------- Settings ----------
$SourceSheetName = 'EmpMonthTotals'
$TargetSheetName = 'Timesheet Hourly'
$RegularRows = @(7,8,9,10,11)  # 5 consecutive rows for regular work, matching the updated template
$RegularLabel = 'יום עבודה רגיל '
$RegularNetwork = 'M01192'
$RegularActivity = '1001'

# Source columns in EmpMonthTotals.xlsx
$ColExit  = 41  # AO
$ColEntry = 47  # AU
$ColDay   = 57  # BE
$ColMarker = 59 # BG - weekly/monthly total marker

function Select-ExcelFile([string]$Title, [string]$Filter) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = $Title
    $dlg.Filter = $Filter
    $dlg.Multiselect = $false
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $dlg.FileName
}

function Convert-ClockValueToExcelTime($Value) {
    if ($null -eq $Value -or "$Value".Trim() -eq '') { return $null }

    # Already a DateTime (some Excel exports may store real times)
    if ($Value -is [datetime]) {
        return ($Value.Hour * 3600 + $Value.Minute * 60 + $Value.Second) / 86400.0
    }

    $s = "$Value".Trim().Replace(',', '.')

    # Text in HH:mm / H:mm form
    if ($s -match '^(\d{1,2}):(\d{2})$') {
        $h = [int]$matches[1]; $m = [int]$matches[2]
        if ($h -gt 23 -or $m -gt 59) { throw "Invalid clock value: $Value" }
        return ($h * 60 + $m) / 1440.0
    }

    # Soroka report stores e.g. 7.12 = 07:12, 12.55 = 12:55
    if ($s -match '^(\d{1,2})(?:\.(\d{1,2}))?$') {
        $h = [int]$matches[1]
        $m = 0
        if ($matches[2]) { $m = [int]$matches[2].PadRight(2,'0') }
        if ($h -gt 23 -or $m -gt 59) { throw "Invalid clock value: $Value" }
        return ($h * 60 + $m) / 1440.0
    }

    throw "Unrecognized clock value: $Value"
}

function Format-TimeForMessage($ExcelTime) {
    if ($null -eq $ExcelTime) { return '--:--' }
    $mins = [math]::Round([double]$ExcelTime * 1440)
    return ('{0:00}:{1:00}' -f [math]::Floor($mins / 60), ($mins % 60))
}


function Test-TimesheetCompletion {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][int]$Year,
        [Parameter(Mandatory=$true)][int]$Month,
        [Parameter(Mandatory=$true)]$ExcelApp
    )

    $missing = New-Object System.Collections.ArrayList
    $incomplete = New-Object System.Collections.ArrayList
    $checkBook = $null
    $checkSheet = $null

    try {
        # Open a fresh read-only view of the SAVED output file.
        # This means manual corrections in Excel must be saved before pressing R.
        $checkBook = $ExcelApp.Workbooks.Open($Path, 0, $true)
        try { $checkSheet = $checkBook.Worksheets.Item($TargetSheetName) }
        catch { throw "Sheet '$TargetSheetName' was not found during recheck." }

        $daysInMonth = [DateTime]::DaysInMonth($Year, $Month)
        $today = Get-Date
        $checkThroughDay = $daysInMonth
        if ($today.Year -eq $Year -and $today.Month -eq $Month) { $checkThroughDay = $today.Day }

        # Rows 7-47 contain all report types in the supplied template:
        # regular work, vacation, sick leave, admin, etc.
        $firstDataRow = 7
        $lastDataRow = 47

        for ($day = 1; $day -le $checkThroughDay; $day++) {
            $date = Get-Date -Year $Year -Month $Month -Day $day
            if ($date.DayOfWeek -eq [DayOfWeek]::Friday -or $date.DayOfWeek -eq [DayOfWeek]::Saturday) { continue }

            $fromCol = 5 + (($day - 1) * 3)
            $toCol = $fromCol + 1
            $hasCompletePair = $false
            $hasPartialPair = $false

            for ($row = $firstDataRow; $row -le $lastDataRow; $row++) {
                $fromVal = $checkSheet.Cells.Item($row,$fromCol).Value2
                $toVal   = $checkSheet.Cells.Item($row,$toCol).Value2
                $hasFrom = ($null -ne $fromVal -and "$fromVal".Trim() -ne '')
                $hasTo   = ($null -ne $toVal   -and "$toVal".Trim() -ne '')

                if ($hasFrom -and $hasTo) { $hasCompletePair = $true }
                elseif ($hasFrom -or $hasTo) { $hasPartialPair = $true }
            }

            if (-not $hasCompletePair -and -not $hasPartialPair) {
                [void]$missing.Add($date)
            }
            elseif ($hasPartialPair) {
                [void]$incomplete.Add($date)
            }
        }

        return [pscustomobject]@{
            Missing    = $missing
            Incomplete = $incomplete
        }
    }
    finally {
        if ($checkBook) { try { $checkBook.Close($false) } catch {} }
        if ($checkSheet) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($checkSheet) | Out-Null } catch {} }
        if ($checkBook) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($checkBook) | Out-Null } catch {} }
    }
}

# ---------- Pick files ----------
if ([string]::IsNullOrWhiteSpace($AttendanceFile)) {
    $AttendanceFile = Select-ExcelFile 'בחר קובץ נוכחות EmpMonthTotals.xlsx' 'Excel (*.xlsx)|*.xlsx'
    if (-not $AttendanceFile) { exit }
}
if ([string]::IsNullOrWhiteSpace($TemplateFile)) {
    $TemplateFile = Select-ExcelFile 'בחר Timesheet Report Template.xls' 'Excel 97-2003 (*.xls)|*.xls|Excel (*.xlsx)|*.xlsx'
    if (-not $TemplateFile) { exit }
}

$AttendanceFile = [IO.Path]::GetFullPath($AttendanceFile)
$TemplateFile = [IO.Path]::GetFullPath($TemplateFile)

if (-not (Test-Path $AttendanceFile)) { throw "Attendance file not found: $AttendanceFile" }
if (-not (Test-Path $TemplateFile)) { throw "Template file not found: $TemplateFile" }

if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $dir = Split-Path $TemplateFile -Parent
    $OutputFile = Join-Path $dir ('Timesheet_Ready_{0}.xls' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$OutputFile = [IO.Path]::GetFullPath($OutputFile)

$excel = $null
$srcBook = $null
$dstBook = $null
$warnings = New-Object System.Collections.ArrayList
$summary = New-Object System.Collections.ArrayList
$emptyDays = New-Object System.Collections.ArrayList

try {
    Show-RealProgress 5 'Starting Excel...'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # ---------- Read attendance ----------
    Show-RealProgress 10 'Opening attendance file...'
    $srcBook = $excel.Workbooks.Open($AttendanceFile, 0, $true)
    try { $srcSheet = $srcBook.Worksheets.Item($SourceSheetName) }
    catch { $srcSheet = $srcBook.Worksheets.Item(1) }

    # Month/year is stored in A1 as MM/YYYY (e.g. 10/2026)
    $monthText = "$($srcSheet.Cells.Item(1,1).Text)".Trim()
    if ($monthText -notmatch '^(\d{1,2})/(\d{4})$') {
        throw "Could not read month/year from A1. Found: '$monthText'"
    }
    $month = [int]$matches[1]
    $year  = [int]$matches[2]

    # day -> list of intervals. Supports duplicate day rows if future exports contain them.
    $days = @{}
    $lastRow = $srcSheet.UsedRange.Rows.Count

    for ($r = 1; $r -le $lastRow; $r++) {
        $readPct = 15 + [int](30 * ($r / [math]::Max(1,$lastRow)))
        Show-RealProgress $readPct ("Reading attendance row $r of $lastRow")
        $marker = "$($srcSheet.Cells.Item($r,$ColMarker).Text)".Trim()
        if ($marker -match '^ס\.\s*(שבועי|חודשי)') { continue }

        $dayRaw = $srcSheet.Cells.Item($r,$ColDay).Value2
        if ($null -eq $dayRaw -or "$dayRaw".Trim() -eq '') { continue }

        $day = 0
        if (-not [int]::TryParse("$dayRaw", [ref]$day)) { continue }
        if ($day -lt 1 -or $day -gt 31) { continue }

        $entryRaw = $srcSheet.Cells.Item($r,$ColEntry).Value2
        $exitRaw  = $srcSheet.Cells.Item($r,$ColExit).Value2

        $entry = $null; $exit = $null
        if ($null -ne $entryRaw -and "$entryRaw".Trim() -ne '') {
            try { $entry = Convert-ClockValueToExcelTime $entryRaw }
            catch { [void]$warnings.Add([string]("Day ${day}: invalid entry '$entryRaw'")) }
        }
        if ($null -ne $exitRaw -and "$exitRaw".Trim() -ne '') {
            try { $exit = Convert-ClockValueToExcelTime $exitRaw }
            catch { [void]$warnings.Add([string]("Day ${day}: invalid exit '$exitRaw'")) }
        }

        # Empty day = no attendance. Leave template empty for manual vacation/sick handling.
        if ($null -eq $entry -and $null -eq $exit) { continue }

        if (-not $days.ContainsKey($day)) {
            $days[$day] = New-Object System.Collections.ArrayList
        }
        [void]$days[$day].Add([pscustomobject]@{ Entry=$entry; Exit=$exit; SourceRow=$r })
    }

    $srcBook.Close($false)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($srcSheet) | Out-Null
    [Runtime.InteropServices.Marshal]::ReleaseComObject($srcBook) | Out-Null
    $srcBook = $null

    # Find workdays (Sun-Thu) with no attendance at all.
    # For the current month, check only up to today so future dates are not reported as missing.
    $daysInMonth = [DateTime]::DaysInMonth($year, $month)
    $today = Get-Date
    $checkThroughDay = $daysInMonth
    if ($today.Year -eq $year -and $today.Month -eq $month) { $checkThroughDay = $today.Day }
    for ($d = 1; $d -le $checkThroughDay; $d++) {
        $date = Get-Date -Year $year -Month $month -Day $d
        if ($date.DayOfWeek -ne [DayOfWeek]::Friday -and $date.DayOfWeek -ne [DayOfWeek]::Saturday) {
            if (-not $days.ContainsKey($d)) { [void]$emptyDays.Add($date) }
        }
    }

    # ---------- Open target template ----------
    Show-RealProgress 50 'Opening timesheet template...'
    $dstBook = $excel.Workbooks.Open($TemplateFile)
    try { $ws = $dstBook.Worksheets.Item($TargetSheetName) }
    catch { throw "Sheet '$TargetSheetName' was not found in the template." }

    # Set month/year in the template
    $monthNames = @('', 'January','February','March','April','May','June','July','August','September','October','November','December')
    $ws.Range('B2').Value2 = [string]$monthNames[[int]$month]
    $ws.Range('B3').Value2 = [string]$year

    Show-RealProgress 60 'Preparing regular-work rows...'

    # Prepare 5 regular-work rows without inserting rows (keeps template structure stable).
    # Row 7 exists; copy its formulas/format to rows 11-14, while rows 8-10 stay untouched.
    foreach ($row in $RegularRows) {
        if ($row -ne 7) {
            [void]$ws.Rows.Item(7).Copy($ws.Rows.Item($row))
        }
        $ws.Cells.Item($row,1).Value2 = [string]$RegularLabel
        $ws.Cells.Item($row,2).Value2 = [string]$RegularNetwork
        $ws.Cells.Item($row,3).Value2 = [string]$RegularActivity
    }

    # Clear only input cells on regular-work rows for all 31 days.
    # Each day uses 3 cols: FROM, TO, TOTAL. Day 1 starts at E (5).
    for ($day=1; $day -le 31; $day++) {
        $prepPct = 60 + [int](10 * ($day / 31.0))
        Show-RealProgress $prepPct ("Preparing day $day of 31")
        $fromCol = 5 + (($day-1) * 3)
        $toCol = $fromCol + 1
        foreach ($row in $RegularRows) {
            [void]$ws.Cells.Item($row,$fromCol).ClearContents()
            [void]$ws.Cells.Item($row,$toCol).ClearContents()
            $ws.Cells.Item($row,$fromCol).NumberFormat = 'hh:mm'
            $ws.Cells.Item($row,$toCol).NumberFormat = 'hh:mm'
        }
    }

    # Fill attendance
    $dayKeys = @($days.Keys | Sort-Object)
    $dayIndex = 0
    foreach ($day in $dayKeys) {
        $dayIndex++
        $fillPct = 70 + [int](20 * ($dayIndex / [math]::Max(1,$dayKeys.Count)))
        Show-RealProgress $fillPct ("Writing day $day ($dayIndex of $($dayKeys.Count))")
        $intervals = @($days[$day])
        $fromCol = 5 + (($day-1) * 3)
        $toCol = $fromCol + 1

        if ($intervals.Count -gt $RegularRows.Count) {
            [void]$warnings.Add([string]("Day ${day}: $($intervals.Count) intervals found; only first $($RegularRows.Count) were written."))
        }

        for ($i=0; $i -lt [Math]::Min($intervals.Count,$RegularRows.Count); $i++) {
            $it = $intervals[$i]
            $targetRow = $RegularRows[$i]

            if ($null -ne $it.Entry) { $ws.Cells.Item($targetRow,$fromCol).Value2 = [double]$it.Entry }
            if ($null -ne $it.Exit)  { $ws.Cells.Item($targetRow,$toCol).Value2   = [double]$it.Exit }

            if ($null -eq $it.Entry -or $null -eq $it.Exit) {
                [void]$warnings.Add([string]("Day ${day}: incomplete punch - entry $(Format-TimeForMessage $it.Entry), exit $(Format-TimeForMessage $it.Exit). Check manually."))
            }
            elseif ([double]$it.Exit -lt [double]$it.Entry) {
                [void]$warnings.Add([string]("Day ${day}: exit is earlier than entry ($(Format-TimeForMessage $it.Entry)-$(Format-TimeForMessage $it.Exit)). Check manually."))
            }

            [void]$summary.Add([string]("Day ${day}: $(Format-TimeForMessage $it.Entry) - $(Format-TimeForMessage $it.Exit)"))
        }
    }

    Show-RealProgress 95 'Saving finished timesheet...'

    # Save as XLS because the upload site expects the old template format.
    # 56 = xlExcel8 (Excel 97-2003 Workbook *.xls)
    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
    $dstBook.SaveAs($OutputFile, 56)
    $dstBook.Close($true)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
    [Runtime.InteropServices.Marshal]::ReleaseComObject($dstBook) | Out-Null
    $dstBook = $null
    Show-RealProgress 100 'Completed'
    Start-Sleep -Milliseconds 400

    Show-AlonHeader
    Write-Host '  PROCESS COMPLETED' -ForegroundColor Green
    Write-Host ('  ' + ('-' * 68)) -ForegroundColor DarkMagenta
    Write-Host ("  Month              : {0:00}/{1}" -f $month,$year) -ForegroundColor White
    Write-Host ("  Filled days        : {0}" -f $days.Count) -ForegroundColor Green
    Write-Host ("  Warnings           : {0}" -f $warnings.Count) -ForegroundColor Yellow
    Write-Host ("  No attendance      : {0}" -f $emptyDays.Count) -ForegroundColor Red
    Write-Host ''

    if ($summary.Count -gt 0) {
        Write-Host '  [OK] ATTENDANCE COPIED' -ForegroundColor Green
        $summary | ForEach-Object { Write-Host "      $_" -ForegroundColor Gray }
    } else {
        Write-Host '  [WARNING] No attendance intervals were found.' -ForegroundColor Yellow
    }

    Write-Host ''
    if ($emptyDays.Count -gt 0) {
        Write-Host '  [EMPTY] WORKDAYS WITH NO ATTENDANCE' -ForegroundColor Red
        foreach ($emptyDate in $emptyDays) {
            Write-Host ("      {0:dd/MM/yyyy} ({1})" -f $emptyDate, $emptyDate.DayOfWeek) -ForegroundColor Red
        }
    } else {
        Write-Host '  [OK] No empty workdays detected.' -ForegroundColor Green
    }

    Write-Host ''
    if ($warnings.Count -gt 0) {
        Write-Host '  [WARNING] CHECK MANUALLY' -ForegroundColor Yellow
        $warnings | ForEach-Object { Write-Host "      $_" -ForegroundColor Yellow }
    } else {
        Write-Host '  [OK] No punch problems detected.' -ForegroundColor Green
    }

    Write-Host ''
    Write-Host '  Vacation / sick / admin were NOT filled automatically.' -ForegroundColor DarkYellow
    Write-Host '  Review the Excel file before uploading it.' -ForegroundColor DarkYellow
    Write-Host ''
    Write-Host '  OUTPUT FILE' -ForegroundColor Magenta
    Write-Host ("  $OutputFile") -ForegroundColor White
    Write-Host ('  ' + ('-' * 68)) -ForegroundColor DarkMagenta

    # Open the finished file for review and keep the console open.
    Start-Process $OutputFile

    Write-Host ''
    Write-Host '  Excel was opened for review.' -ForegroundColor Magenta
    Write-Host '  Make your manual corrections in Excel and SAVE the file.' -ForegroundColor Gray
    Write-Host ''

    while ($true) {
        Write-Host '  [R] Recheck after corrections' -ForegroundColor Cyan
        Write-Host '  [Q] Finish' -ForegroundColor DarkGray
        Write-Host '  Select: ' -NoNewline -ForegroundColor White
        $pressedKey = [Console]::ReadKey($true).Key

        if ($pressedKey -eq [ConsoleKey]::Q) {
            Write-Host 'Q' -ForegroundColor DarkGray
            break
        }

        if ($pressedKey -ne [ConsoleKey]::R) {
            Write-Host '?' -ForegroundColor Yellow
            Write-Host '  Please press R or Q.' -ForegroundColor Yellow
            Write-Host ''
            continue
        }

        Write-Host 'R' -ForegroundColor Cyan

        Show-AlonHeader
        Write-Host '  RECHECKING SAVED TIMESHEET...' -ForegroundColor Cyan
        Write-Host '  Reading the output file after your manual corrections.' -ForegroundColor Gray
        Write-Host ''

        try {
            $recheck = Test-TimesheetCompletion -Path $OutputFile -Year $year -Month $month -ExcelApp $excel

            Write-Host '  RECHECK RESULTS' -ForegroundColor Magenta
            Write-Host ('  ' + ('-' * 68)) -ForegroundColor DarkMagenta

            if ($recheck.Missing.Count -eq 0 -and $recheck.Incomplete.Count -eq 0) {
                Write-Host ''
                Write-Host '  [OK] ALL REQUIRED WORK DAYS ARE FILLED' -ForegroundColor Green
                Write-Host '  No missing or partial attendance was found.' -ForegroundColor Green
                Write-Host '  Timesheet is ready for your final review/upload.' -ForegroundColor White
                Write-Host ''
                [void](Read-Host '  Press Enter to finish')
                break
            }

            if ($recheck.Missing.Count -gt 0) {
                Write-Host ''
                Write-Host '  [EMPTY] DAYS WITH NO DATA' -ForegroundColor Red
                foreach ($d in $recheck.Missing) {
                    Write-Host ("      {0:dd/MM/yyyy} ({1})" -f $d,$d.DayOfWeek) -ForegroundColor Red
                }
            }

            if ($recheck.Incomplete.Count -gt 0) {
                Write-Host ''
                Write-Host '  [WARNING] DAYS WITH PARTIAL FROM/TO DATA' -ForegroundColor Yellow
                foreach ($d in $recheck.Incomplete) {
                    Write-Host ("      {0:dd/MM/yyyy} ({1})" -f $d,$d.DayOfWeek) -ForegroundColor Yellow
                }
            }

            Write-Host ''
            Write-Host ("  Missing days    : {0}" -f $recheck.Missing.Count) -ForegroundColor Red
            Write-Host ("  Partial days    : {0}" -f $recheck.Incomplete.Count) -ForegroundColor Yellow
            Write-Host ''
            Write-Host '  Return to Excel, correct the report and SAVE it.' -ForegroundColor Gray
            Write-Host '  Then press R again to run another check.' -ForegroundColor Gray
            Write-Host ''
        }
        catch {
            Write-Host ''
            Write-Host ("  RECHECK ERROR: {0}" -f $_.Exception.Message) -ForegroundColor Red
            Write-Host '  Make sure the Excel file is saved, then try R again.' -ForegroundColor Yellow
            Write-Host ''
        }
    }
}
catch {
    Write-Host ''
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkYellow
    Write-Host "Command: $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkYellow
    Write-Host 'No file should be uploaded until this error is resolved.' -ForegroundColor Red
    Read-Host 'Press Enter to close'
}
finally {
    if ($srcBook) { try { $srcBook.Close($false) } catch {} }
    if ($dstBook) { try { $dstBook.Close($false) } catch {} }
    if ($excel) {
        try { $excel.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
