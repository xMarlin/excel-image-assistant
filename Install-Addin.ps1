$ErrorActionPreference = 'Stop'
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Source = Join-Path $Here 'Excel Image Assistant - Mohamed v1.7.xlam'
if (-not (Test-Path $Source)) { throw "Build the add-in first. Missing: $Source" }

$AddinsDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
New-Item -ItemType Directory -Path $AddinsDir -Force | Out-Null
$Dest = Join-Path $AddinsDir 'Excel Image Assistant - Mohamed v1.7.xlam'
Copy-Item $Source $Dest -Force
try { Unblock-File -Path $Dest -ErrorAction SilentlyContinue } catch { }

$excel=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false
    $excel.DisplayAlerts=$false

    # Disable only older Mohamed clone registrations; do not touch the vendor/original add-in.
    foreach($ai in @($excel.AddIns)) {
        try {
            if($ai.FullName -like '*Excel Image Assistant - Mohamed*' -and $ai.FullName -ne $Dest) {
                $ai.Installed=$false
            }
        } catch { }
    }

    $addin=$null
    foreach($ai in @($excel.AddIns)) {
        try { if([string]::Equals($ai.FullName,$Dest,[StringComparison]::OrdinalIgnoreCase)){ $addin=$ai; break } } catch { }
    }
    if($null -eq $addin){ $addin=$excel.AddIns.Add($Dest,$true) }
    $addin.Installed=$true
    $excel.Quit(); $excel=$null
} finally {
    if($null -ne $excel){ try{$excel.Quit()}catch{} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Host ''
Write-Host 'INSTALLED: Excel Image Assistant - Mohamed v1.7' -ForegroundColor Green
Write-Host 'Close every Excel window, then reopen Excel.' -ForegroundColor Yellow
Write-Host 'The visible tab name is: Excel Image Assistant' -ForegroundColor Cyan
