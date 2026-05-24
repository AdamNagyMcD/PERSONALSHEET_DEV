$base = "c:\DEV\PERSONALSHEET_DEV\_xlsm_user_ref2"
$sheetPath = "$base\xl\worksheets\sheet3.xml"
[xml]$styles = Get-Content "$base\xl\styles.xml" -Encoding UTF8
[xml]$sheet = Get-Content $sheetPath -Encoding UTF8

function ClassifyFill([string]$rgb) {
    if (-not $rgb) { return "default" }
    switch -Regex ($rgb.ToUpper()) {
        "FFDDEBF7|FFDCE6F1" { return "HEADER" }
        "FFFFF2CC|FFFFC000|FFFFEB9C|FFFFD966" { return "ACCENT" }
        "FFFFFFFF" { return "WHITE" }
        "FFF8F8F8|FFF2F2F2|FFF5F5F5" { return "ZEBRA" }
        "FF1F4E79|FF4472C4" { return "NAVY" }
        default { return "OTHER:$rgb" }
    }
}

$fills = @($null)
foreach ($f in $styles.styleSheet.fills.fill) {
    if ($f.patternFill.fgColor.rgb) { $fills += $f.patternFill.fgColor.rgb }
    else { $fills += "none" }
}
$xfsFill = @()
foreach ($xf in $styles.styleSheet.cellXfs.xf) { $xfsFill += [int]$xf.fillId }

$cells = @()
foreach ($row in $sheet.worksheet.sheetData.row) {
    foreach ($c in $row.c) {
        if (-not $c.r) { continue }
        if ($c.r -match '^([A-Z]+)(\d+)$') {
            $colLetters = $matches[1]
            $rowNum = [int]$matches[2]
            $colNum = 0
            foreach ($ch in $colLetters.ToCharArray()) { $colNum = $colNum * 26 + ([int][char]$ch - [int][char]'A' + 1) }
            $s = 0; if ($c.s) { $s = [int]$c.s }
            $rgb = $fills[$xfsFill[$s] + 1]
            $cells += [PSCustomObject]@{ Ref = $c.r; Row = $rowNum; Col = $colNum; ColL = $colLetters; Role = (ClassifyFill $rgb); Rgb = $rgb }
        }
    }
}

Write-Output "=== ROW 8-16 (crew stats) actual cells ==="
$cells | Where-Object { $_.Row -ge 8 -and $_.Row -le 16 -and $_.Col -ge 15 -and $_.Col -le 18 } | Sort-Object Row, Col | Format-Table Row, Ref, Role -AutoSize

Write-Output "=== ROW 17-29 actual cells O-R ==="
$cells | Where-Object { $_.Row -ge 17 -and $_.Row -le 29 -and $_.Col -ge 15 -and $_.Col -le 18 } | Sort-Object Row, Col | Format-Table Row, Ref, Role -AutoSize

Write-Output "=== ROW 31-45 actual cells O-V ==="
$cells | Where-Object { $_.Row -ge 31 -and $_.Row -le 45 -and $_.Col -ge 15 -and $_.Col -le 22 } | Sort-Object Row, Col | Format-Table Row, Ref, Role -AutoSize

Write-Output "=== EMPLOYEE cols B-N rows 1-3 sample ==="
$cells | Where-Object { $_.Row -le 3 -and $_.Col -ge 2 -and $_.Col -le 14 } | Sort-Object Row, Col | Format-Table Row, Ref, Role -AutoSize

# Export rules as JSON for implementation
$rules = @{
    employee = @{}
    panel = @{}
}
foreach ($r in 3..82) {
    $rowCells = $cells | Where-Object { $_.Row -eq $r -and $_.Col -ge 2 -and $_.Col -le 14 }
    if ($rowCells) {
        $roles = ($rowCells | Group-Object Role | ForEach-Object { "$($_.Name)" }) -join "/"
        $rules.employee["row$r"] = $roles
    }
}
foreach ($r in 3..50) {
    $rowCells = $cells | Where-Object { $_.Row -eq $r -and $_.Col -ge 15 -and $_.Col -le 22 }
    if ($rowCells) {
        $detail = $rowCells | Sort-Object Col | ForEach-Object { "$($_.ColL):$($_.Role)" }
        $rules.panel["row$r"] = ($detail -join ", ")
    }
}
$rules | ConvertTo-Json -Depth 4 | Set-Content "c:\DEV\PERSONALSHEET_DEV\tools\januar_style_rules.json" -Encoding UTF8
Write-Output "Rules written to tools/januar_style_rules.json"
