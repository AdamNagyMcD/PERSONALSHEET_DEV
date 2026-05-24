$base = "c:\DEV\PERSONALSHEET_DEV\_xlsm_user_ref2"
$sheetPath = "$base\xl\worksheets\sheet3.xml"

[xml]$styles = Get-Content "$base\xl\styles.xml" -Encoding UTF8
[xml]$sheet = Get-Content $sheetPath -Encoding UTF8

function ColToNum([string]$col) {
    $n = 0
    foreach ($ch in $col.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1) }
    return $n
}

function NumToCol([int]$n) {
    $s = ""
    while ($n -gt 0) {
        $m = ($n - 1) % 26
        $s = [char]([int][char]'A' + $m) + $s
        $n = [int](($n - 1) / 26)
    }
    return $s
}

function ClassifyFill([string]$rgb) {
    if (-not $rgb) { return "default" }
    switch -Regex ($rgb.ToUpper()) {
        "FFDDEBF7|FFDCE6F1" { return "HEADER" }
        "FFFFF2CC|FFFFC000|FFFFEB9C|FFFFD966" { return "ACCENT" }
        "FFFFFFFF" { return "WHITE" }
        "FFF8F8F8|FFE7E6E6|FFD9D9D9" { return "ZEBRA" }
        "FF1F4E79|FF4472C4" { return "NAVY" }
        default { return $rgb }
    }
}

$fills = @($null)
foreach ($f in $styles.styleSheet.fills.fill) {
    $rgb = $null
    if ($f.patternFill.fgColor.rgb) { $rgb = $f.patternFill.fgColor.rgb }
    elseif ($f.patternFill.bgColor.rgb) { $rgb = $f.patternFill.bgColor.rgb }
    else { $rgb = "none" }
    $fills += $rgb
}

$xfsFill = @()
foreach ($xf in $styles.styleSheet.cellXfs.xf) { $xfsFill += [int]$xf.fillId }

$cellMap = @{}
foreach ($row in $sheet.worksheet.sheetData.row) {
    $r = [int]$row.r
    foreach ($c in $row.c) {
        if (-not $c.r) { continue }
        if ($c.r -match '^([A-Z]+)(\d+)$') {
            $colNum = ColToNum $matches[1]
            $rowNum = [int]$matches[2]
            $styleIdx = 0
            if ($c.s) { $styleIdx = [int]$c.s }
            $fillId = $xfsFill[$styleIdx]
            $rgb = $fills[$fillId + 1]
            $cellMap["$colNum|$rowNum"] = @{ Ref = $c.r; Row = $rowNum; Col = $colNum; Fill = $rgb; Role = (ClassifyFill $rgb) }
        }
    }
}

function Get-Role([int]$col, [int]$row) {
    $k = "$col|$row"
    if ($cellMap.ContainsKey($k)) { return $cellMap[$k].Role }
    return "."
}

Write-Output "=== EMPLOYEE TABLE B-N rows 1-15, cols sample ==="
foreach ($r in 1..15) {
    $parts = @()
    foreach ($col in 2,5,6,7,8,11,12) {
        $parts += ("{0}:{1}" -f (NumToCol $col), (Get-Role $col $r))
    }
    Write-Output ("Row {0,2}: {1}" -f $r, ($parts -join "  "))
}

Write-Output ""
Write-Output "=== RIGHT PANEL O-V rows 3-50 by row (dominant roles) ==="
foreach ($r in 3..50) {
    $roles = @{}
    foreach ($col in 15..22) {
        $role = Get-Role $col $r
        if ($role -ne ".") {
            if (-not $roles.ContainsKey($role)) { $roles[$role] = 0 }
            $roles[$role]++
        }
    }
    if ($roles.Count -eq 0) { continue }
    $summary = ($roles.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
    Write-Output ("Row {0,2}: {1}" -f $r, $summary)
}

Write-Output ""
Write-Output "=== DETAIL rows 8-29 cols O-R ==="
foreach ($r in 8..29) {
    $parts = @()
    foreach ($col in 15..18) {
        $parts += ("{0}:{1}" -f (NumToCol $col), (Get-Role $col $r))
    }
    Write-Output ("Row {0,2}: {1}" -f $r, ($parts -join "  "))
}

Write-Output ""
Write-Output "=== DETAIL rows 33-45 cols O-V ==="
foreach ($r in 33..45) {
    $parts = @()
    foreach ($col in 15..22) {
        $parts += ("{0}:{1}" -f (NumToCol $col), (Get-Role $col $r))
    }
    Write-Output ("Row {0,2}: {1}" -f $r, ($parts -join "  "))
}

Write-Output ""
Write-Output "=== UNIQUE FILL RGBs used on sheet ==="
$cellMap.Values | Group-Object Fill | Sort-Object Count -Descending | Select-Object -First 15 | ForEach-Object {
    "{0}  ({1})  count={2}" -f (ClassifyFill $_.Name), $_.Name, $_.Count
}
