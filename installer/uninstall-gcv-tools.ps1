# Removes GCV Tools for the current Windows user. Excel must be closed.
# Logs (%LOCALAPPDATA%\GCV\ExcelLogs) are kept.

$ErrorActionPreference = 'Stop'
$officeVersion = '16.0'
$base = Join-Path $env:LOCALAPPDATA 'GCV\ExcelTools'
$excelKey = "HKCU:\Software\Microsoft\Office\$officeVersion\Excel"
$optionsKey = "$excelKey\Options"

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host 'ERROR: Excel is open. Close every Excel window and try again.' -ForegroundColor Red
    exit 1
}
if (Test-Path $optionsKey) {
    $options = Get-ItemProperty -Path $optionsKey
    $options.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' -and $_.Value -like '*GCV Loader.xlam*' } |
        ForEach-Object { Remove-ItemProperty -Path $optionsKey -Name $_.Name; Write-Host "  removed $($_.Name)" }
}
Remove-Item -Path "$excelKey\Security\Trusted Locations\GCVTools" -Recurse -ErrorAction SilentlyContinue
Remove-Item -Path 'HKCU:\Software\VB and VBA Program Settings\GCV Tools' -Recurse -ErrorAction SilentlyContinue
if (Test-Path $base) { Remove-Item -LiteralPath $base -Recurse -Force }
Write-Host 'GCV Tools removed.' -ForegroundColor Green
