# Removes GCV Tools for the current Windows user. Excel must be closed.
# Logs (%LOCALAPPDATA%\GCV\ExcelLogs) are kept.

$ErrorActionPreference = 'Stop'
$officeVersion = '16.0'
$base = Join-Path $env:LOCALAPPDATA 'GCV\ExcelTools'
$excelKey = "HKCU:\Software\Microsoft\Office\$officeVersion\Excel"
$optionsKey = "$excelKey\Options"

function Fail([string]$message) {
    Write-Host ''
    Write-Host "ERROR: $message" -ForegroundColor Red
    exit 1
}

# Excel of THIS Windows session only (Get-Process alone also lists Excel of other
# signed-in users). A hidden Excel without a window (left after a crash or
# automation) can be closed here; one with a window must be closed by the user.
function Confirm-ExcelClosed {
    $session = (Get-Process -Id $PID).SessionId
    $excel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $session })
    if ($excel.Count -eq 0) { return }
    $visible = @($excel | Where-Object { $_.MainWindowHandle -ne 0 })
    if ($visible.Count -gt 0) {
        Fail "Excel is open ($($visible[0].MainWindowTitle)). Close every Excel window and run this again."
    }
    Write-Host ''
    Write-Host "Excel is still running in the background without a window (process $(($excel | ForEach-Object { $_.Id }) -join ', '))." -ForegroundColor Yellow
    $answer = Read-Host 'Close it now? Unsaved work in it would be lost. [Y/N]'
    if ($answer -notmatch '^[Yy]$') {
        Fail 'Excel is still running. Close it in Task Manager (Details tab, EXCEL.EXE) and run this again.'
    }
    $excel | Stop-Process -Force
    Start-Sleep -Seconds 2
    Write-Host '  background Excel closed'
}

Confirm-ExcelClosed
if (Test-Path $optionsKey) {
    $options = Get-ItemProperty -Path $optionsKey
    $options.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' -and $_.Value -like '*GCV Loader.xlam*' } |
        ForEach-Object { Remove-ItemProperty -Path $optionsKey -Name $_.Name; Write-Host "  removed $($_.Name)" }
}
Remove-Item -Path "$excelKey\Security\Trusted Locations\GCVTools" -Recurse -ErrorAction SilentlyContinue
Remove-Item -Path 'HKCU:\Software\VB and VBA Program Settings\GCV Tools' -Recurse -ErrorAction SilentlyContinue
if (Test-Path $base) { Remove-Item -LiteralPath $base -Recurse -Force }
Write-Host 'GCV Tools removed.' -ForegroundColor Green
