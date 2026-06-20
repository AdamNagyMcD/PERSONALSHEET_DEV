# READ-ONLY Layout-Probe fuer das FLUKTUATION-Blatt.
# Liest die aktuell gespeicherte Formatierung aus und schreibt sie als JSON-Bericht.
#
# Garantiert read-only:
#   - oeffnet die Mappe mit ReadOnly:=True
#   - EnableEvents=False (Workbook_Open/Refresh laeuft NICHT)
#   - KEIN Save, KEIN Import, KEINE VBA-Aenderung
#   - schliesst mit Close($false)
# Voraussetzung: Excel ist geschlossen (Mappe nicht gesperrt).

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$wbPath = Join-Path $root "Personalsheet.xlsm"
$outPath = Join-Path $PSScriptRoot "fluktuation_styles.json"
$sheetName = "FLUKTUATION"

# Maximaler Auslesebereich (Spalten A..V, Zeilen 1..<lastRow>, gedeckelt).
$maxCol = 22
$maxRowCap = 90

function ToRgbHex($oleColor) {
    if ($null -eq $oleColor) { return $null }
    $v = [int64]$oleColor
    if ($v -lt 0) { return "auto" }   # xlColorIndexAutomatic o.ae.
    $r = $v -band 0xFF
    $g = ($v -shr 8) -band 0xFF
    $b = ($v -shr 16) -band 0xFF
    return ("{0:X2}{1:X2}{2:X2}" -f $r, $g, $b)
}

# Lock-Check.
try {
    $fs = [System.IO.File]::Open($wbPath, 'Open', 'ReadWrite', 'None')
    $fs.Close(); $fs.Dispose()
} catch {
    Write-Host "LOCKED: Personalsheet.xlsm ist geoeffnet/gesperrt. Bitte Excel komplett schliessen und erneut ausfuehren."
    exit 2
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.ScreenUpdating = $false
$xl.AutomationSecurity = 3  # ForceDisable: keine Makros beim Lesen

$wb = $null
try {
    $wb = $xl.Workbooks.Open($wbPath, 0, $true)  # UpdateLinks=0, ReadOnly=$true
    $ws = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq $sheetName) { $ws = $s; break } }
    if ($null -eq $ws) { throw "Sheet '$sheetName' not found." }

    $used = $ws.UsedRange
    $lastRow = [int]($used.Row + $used.Rows.Count - 1)
    $lastCol = [int]($used.Column + $used.Columns.Count - 1)
    if ($lastRow -gt $maxRowCap) { $lastRow = $maxRowCap }
    if ($lastCol -gt $maxCol) { $lastCol = $maxCol }

    $report = [ordered]@{
        sheet      = $sheetName
        scannedAt  = (Get-Date).ToString("s")
        lastRow    = $lastRow
        lastCol    = $lastCol
        columns    = @()
        rows       = @()
        merges     = @()
        cells      = @()
    }

    for ($c = 1; $c -le $lastCol; $c++) {
        $report.columns += [ordered]@{ col = $c; letter = [string]($ws.Cells.Item(1, $c).Address($false, $false) -replace '\d', ''); width = [math]::Round([double]$ws.Columns.Item($c).ColumnWidth, 2) }
    }
    for ($r = 1; $r -le $lastRow; $r++) {
        $report.rows += [ordered]@{ row = $r; height = [math]::Round([double]$ws.Rows.Item($r).RowHeight, 2) }
    }

    $mergeSeen = @{}
    for ($r = 1; $r -le $lastRow; $r++) {
        for ($c = 1; $c -le $lastCol; $c++) {
            $cell = $ws.Cells.Item($r, $c)

            if ($cell.MergeCells) {
                $ma = $cell.MergeArea.Address($false, $false)
                if (-not $mergeSeen.ContainsKey($ma)) { $mergeSeen[$ma] = $true; $report.merges += $ma }
            }

            $val = $cell.Value2
            $hasText = ($null -ne $val -and "$val".Trim().Length -gt 0)
            $fill = $null
            try { if ($cell.Interior.Pattern -ne -4142) { $fill = ToRgbHex $cell.Interior.Color } } catch {}
            $bold = [bool]$cell.Font.Bold
            $italic = [bool]$cell.Font.Italic

            # Nur "interessante" Zellen aufnehmen (Inhalt, Fuellung, Fett/Kursiv, Merge-Anker).
            if (-not ($hasText -or $fill -or $bold -or $italic -or $cell.MergeCells)) { continue }

            $b = $cell.Borders
            function EdgeInfo($edge) {
                $w = $null; $st = $null; $col = $null
                try { $w = [int]$edge.Weight } catch {}
                try { $st = [int]$edge.LineStyle } catch {}
                try { $col = ToRgbHex $edge.Color } catch {}
                return [ordered]@{ weight = $w; lineStyle = $st; color = $col }
            }

            $report.cells += [ordered]@{
                addr     = $cell.Address($false, $false)
                row      = $r
                col      = $c
                value    = if ($hasText) { "$val" } else { $null }
                merged   = [bool]$cell.MergeCells
                font     = [ordered]@{
                    name  = [string]$cell.Font.Name
                    size  = [double]$cell.Font.Size
                    bold  = $bold
                    italic= $italic
                    color = (ToRgbHex $cell.Font.Color)
                }
                fill     = $fill
                hAlign   = [int]$cell.HorizontalAlignment
                vAlign   = [int]$cell.VerticalAlignment
                numFmt   = [string]$cell.NumberFormat
                wrap     = [bool]$cell.WrapText
                locked   = [bool]$cell.Locked
                borders  = [ordered]@{
                    left   = (EdgeInfo $b.Item(7))   # xlEdgeLeft
                    top    = (EdgeInfo $b.Item(8))    # xlEdgeTop
                    bottom = (EdgeInfo $b.Item(9))    # xlEdgeBottom
                    right  = (EdgeInfo $b.Item(10))   # xlEdgeRight
                }
            }
        }
    }

    $report | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $outPath -Encoding UTF8

    Write-Host "PROBE_OK"
    Write-Host ("Sheet: {0}  lastRow={1} lastCol={2}" -f $sheetName, $lastRow, $lastCol)
    Write-Host ("Cells erfasst: {0}" -f $report.cells.Count)
    Write-Host ("Merges: {0}" -f $report.merges.Count)
    Write-Host ("JSON: {0}" -f $outPath)
    Write-Host ("ReadOnly geoeffnet: {0}" -f $wb.ReadOnly)
}
finally {
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
