# Imports vba/*.bas + DieseArbeitsmappe.cls into Personalsheet.xlsm, then restores E/F dropdowns.
# Requires: Excel closed (or file not locked) + "Trust access to the VBA project object model"
param(
    [string]$WorkbookPath = (Join-Path (Split-Path $PSScriptRoot -Parent) "Personalsheet.xlsm")
)

$ErrorActionPreference = "Stop"
$vbaFolder = Join-Path (Split-Path $PSScriptRoot -Parent) "vba"

function Read-VbaCodeWithoutHeader {
    param([string]$FullPath)
    $found = $false
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($line in [System.IO.File]::ReadAllLines($FullPath)) {
        $trim = $line.Trim()
        if (-not $found) {
            if ($trim -eq "Option Explicit") {
                $found = $true
                $lines.Add($line)
            }
            continue
        }
        $lines.Add($line)
    }
    if (-not $found) { throw "No Option Explicit in $FullPath" }
    return ($lines -join [Environment]::NewLine)
}

if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
if (-not (Test-Path $vbaFolder)) { throw "VBA folder not found: $vbaFolder" }

try {
    $fs = [System.IO.File]::Open($WorkbookPath, "Open", "ReadWrite", "None")
    $fs.Close()
} catch {
    throw "Personalsheet.xlsm is locked. Close Excel completely, then run this script again."
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null

try {
    $wb = $xl.Workbooks.Open($WorkbookPath, $null, $false)
    $vbProj = $wb.VBProject

    for ($i = $vbProj.VBComponents.Count; $i -ge 1; $i--) {
        $comp = $vbProj.VBComponents.Item($i)
        if ($comp.Type -in 1, 2, 3) {
            if ($comp.Name -ne "mod_ResetAndImportVBAFiles") {
                $vbProj.VBComponents.Remove($comp) | Out-Null
            }
        }
    }

    Get-ChildItem (Join-Path $vbaFolder "*.bas") | Sort-Object Name | ForEach-Object {
        if ($_.Name -ieq "mod_ResetAndImportVBAFiles.bas") { return }
        $vbProj.VBComponents.Import($_.FullName) | Out-Null
        Write-Host "Imported $($_.Name)"
    }

    $clsPath = Join-Path $vbaFolder "DieseArbeitsmappe.cls"
    if (Test-Path $clsPath) {
        $code = Read-VbaCodeWithoutHeader $clsPath
        $cm = $wb.VBProject.VBComponents.Item("DieseArbeitsmappe").CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.InsertLines(1, $code)
        Write-Host "Updated DieseArbeitsmappe.cls"
    }

    $wb.Save()
    Write-Host "VBA saved."

    try {
        $xl.Run("PID_RestoreMonthSheetDropdownsAfterFormatSilent")
        Write-Host "PID_RestoreMonthSheetDropdownsAfterFormatSilent OK"
    } catch {
        Write-Warning "PID_RestoreMonthSheetDropdownsAfterFormatSilent failed: $($_.Exception.Message)"
    }

    try {
        $xl.Run("PID_RestoreAktuelleStundenFormulasSilent")
        Write-Host "PID_RestoreAktuelleStundenFormulasSilent OK"
    } catch {
        Write-Warning "PID_RestoreAktuelleStundenFormulasSilent failed: $($_.Exception.Message)"
    }

    $ws = $wb.Worksheets.Item("Februar")
    $fCount = 0
    3..82 | ForEach-Object {
        try { if ($ws.Cells.Item($_, 6).Validation.Type -eq 3) { $fCount++ } } catch {}
    }
    Write-Host "Februar F dropdown rows: $fCount / 80"

    $wb.Save()
    Write-Host "Done. Saved $WorkbookPath"
}
finally {
    if ($wb) { $wb.Close($true) | Out-Null }
    $xl.Quit() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}
