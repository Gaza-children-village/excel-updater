# Installs GCV Tools for the current Windows user (no administrator rights).
#   1. copies "GCV Loader.xlam" to %LOCALAPPDATA%\GCV\ExcelTools
#      and every other "GCV *.xlam" next to this script (GCV Logger, GCV Updater,
#      GCV Student Lookup, ...) to %LOCALAPPDATA%\GCV\ExcelTools\addins
#   2. makes that folder a trusted location (macros run without warnings)
#   3. registers GCV Loader as an Excel add-in (opens with Excel)
# Run with install-gcv-tools.bat. Excel must be closed. Safe to run again.

$ErrorActionPreference = 'Stop'
$officeVersion = '16.0'                     # Office 2016, 2019, 2021 and Microsoft 365
$source = $PSScriptRoot
$base = Join-Path $env:LOCALAPPDATA 'GCV\ExcelTools'
$addins = Join-Path $base 'addins'
$loaderName = 'GCV Loader.xlam'
$loaderPath = Join-Path $base $loaderName
$excelKey = "HKCU:\Software\Microsoft\Office\$officeVersion\Excel"
$trustKey = "$excelKey\Security\Trusted Locations\GCVTools"
$optionsKey = "$excelKey\Options"

function Fail([string]$message) {
    Write-Host ''
    Write-Host "ERROR: $message" -ForegroundColor Red
    exit 1
}

Write-Host 'GCV Tools - installation' -ForegroundColor Cyan

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Fail 'Excel is open. Close every Excel window and run the installer again.'
}
if (-not (Test-Path "HKCU:\Software\Microsoft\Office\$officeVersion")) {
    Fail "Office $officeVersion was not found for this user. Open Excel once, close it, and try again."
}

$loaderSource = Join-Path $source $loaderName
if (-not (Test-Path -LiteralPath $loaderSource)) { Fail "'$loaderName' is missing next to the installer." }
$toolsFiles = @(Get-ChildItem -LiteralPath $source -Filter 'GCV *.xlam' |
    Where-Object { $_.Name -ne $loaderName -and $_.Name -notmatch 'shell' })
if ($toolsFiles.Count -eq 0) { Fail "No GCV add-ins (GCV Logger.xlam, GCV Updater.xlam, ...) next to the installer." }
if (-not ($toolsFiles.Name -contains 'GCV Logger.xlam') -or -not ($toolsFiles.Name -contains 'GCV Updater.xlam')) {
    Fail "'GCV Logger.xlam' and 'GCV Updater.xlam' are required next to the installer."
}

# 1. Files
New-Item -ItemType Directory -Force -Path $addins | Out-Null
Copy-Item -LiteralPath $loaderSource -Destination $loaderPath -Force
Unblock-File -LiteralPath $loaderPath
foreach ($file in $toolsFiles) {
    $target = Join-Path $addins $file.Name
    Copy-Item -LiteralPath $file.FullName -Destination $target -Force
    Unblock-File -LiteralPath $target
    Write-Host "  copied $($file.Name)"
}
Write-Host "  copied $loaderName"

# 2. Trusted location (per user)
New-Item -Force -Path $trustKey | Out-Null
New-ItemProperty -Path $trustKey -Name 'Path' -Value ($base + '\') -PropertyType String -Force | Out-Null
New-ItemProperty -Path $trustKey -Name 'AllowSubfolders' -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $trustKey -Name 'Description' -Value 'GCV Tools (Gaza Children Village)' -PropertyType String -Force | Out-Null
Write-Host "  trusted location: $base"

# 3. Open the loader with Excel: values OPEN, OPEN1, OPEN2, ... (first free one)
New-Item -Force -Path $optionsKey | Out-Null
$options = Get-ItemProperty -Path $optionsKey
$openNames = @($options.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' })
$already = $openNames | Where-Object { $_.Value -like "*$loaderName*" }
if ($already) {
    foreach ($entry in $already) {
        Set-ItemProperty -Path $optionsKey -Name $entry.Name -Value ('"' + $loaderPath + '"')
    }
    Write-Host '  add-in already registered (path refreshed)'
} else {
    $name = 'OPEN'
    $i = 0
    while ($openNames.Name -contains $name) { $i++; $name = "OPEN$i" }
    New-ItemProperty -Path $optionsKey -Name $name -Value ('"' + $loaderPath + '"') -PropertyType String -Force | Out-Null
    Write-Host "  add-in registered ($name)"
}

Write-Host ''
Write-Host 'Done. Open Excel: the tab "Gaza Children Village Tools" appears after a few seconds.' -ForegroundColor Green
