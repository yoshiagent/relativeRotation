param(
    [string]$Path = ""
)

$ErrorActionPreference = "SilentlyContinue"

function Test-FileLocked {
    param([string]$p)
    if (-not (Test-Path -LiteralPath $p)) { return $false }
    try {
        $f = [System.IO.File]::Open($p, 'Open', 'Read', 'None')
        $f.Close()
        $f.Dispose()
        return $false
    } catch {
        return $true
    }
}

# 若沒傳 -Path，自動從腳本上層目錄找 .xlsx
if (-not $Path) {
    $scriptRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
    $xlsx = Get-ChildItem -LiteralPath $scriptRoot -Filter "*.xlsx" -File | Select-Object -First 1
    if (-not $xlsx) {
        Write-Host "  [skip] No .xlsx found in $scriptRoot"
        exit 0
    }
    $Path = $xlsx.FullName
}

$Target = $Path
Write-Host "  Target: $Target"

if (-not (Test-FileLocked $Target)) {
    Write-Host "  [ok] File is free, no action needed."
    exit 0
}

Write-Host "  File is locked. Attempting graceful close via COM..."

$excel = $null
try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = $null
}

if ($excel) {
    $excel.DisplayAlerts = $false
    $wbList = @($excel.Workbooks)
    foreach ($wb in $wbList) {
        $wbPath = $wb.FullName
        if ($wbPath -ieq $Target) {
            $wbName = $wb.Name
            Write-Host "  Found workbook, saving + closing: $wbName"
            $wb.Save()
            $wb.Close($false)
        }
    }
    if ($excel.Workbooks.Count -eq 0) {
        $excel.Quit()
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
} else {
    Write-Host "  (No running Excel found via COM)"
}

Start-Sleep -Seconds 2

if (Test-FileLocked $Target) {
    Write-Host "  Still locked. Force-killing all EXCEL.EXE..."
    Get-Process -Name "EXCEL" | Stop-Process -Force
    Start-Sleep -Seconds 2
}

if (Test-FileLocked $Target) {
    Write-Host "  [ERROR] Cannot release file lock."
    exit 1
}

Write-Host "  [ok] File released."
exit 0
