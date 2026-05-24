$base = "c:\DEV\PERSONALSHEET_DEV\_xlsm_user_ref2"
$sheetPath = "$base\xl\worksheets\sheet3.xml"
[xml]$styles = Get-Content "$base\xl\styles.xml" -Encoding UTF8
[xml]$sheet = Get-Content $sheetPath -Encoding UTF8

function ColToNum([string]$col) { $n=0; foreach($ch in $col.ToCharArray()){$n=$n*26+[int][char]$ch-[int][char]'A'+1}; return $n }
function ClassifyFill([string]$rgb) {
    if (-not $rgb) { return "default" }
    switch -Regex ($rgb.ToUpper()) {
        "FFDDEBF7|FFDCE6F1" { return "HEADER" }
        "FFFFF2CC|FFFFC000|FFFFEB9C|FFFFD966" { return "ACCENT" }
        "FFFFFFFF" { return "WHITE" }
        "FFF8F8F8" { return "ZEBRA" }
        default { return $rgb }
    }
}
$fills=@($null); foreach($f in $styles.styleSheet.fills.fill){ if($f.patternFill.fgColor.rgb){$fills+=$f.patternFill.fgColor.rgb}else{$fills+="none"} }
$xfsFill=@(); foreach($xf in $styles.styleSheet.cellXfs.xf){$xfsFill+=[int]$xf.fillId}
function GetRole($ref) {
    foreach($row in $sheet.worksheet.sheetData.row){ foreach($c in $row.c){ if($c.r -eq $ref){ $s=0; if($c.s){$s=[int]$c.s}; return (ClassifyFill $fills[$xfsFill[$s]+1]) } } }
    return "."
}

Write-Output "Row 8-16 O-R cell by cell:"
foreach($r in 8..16){ $line=@(); foreach($col in 15..18){ $colL=[char]([int][char]'A'+$col-1); if($col -gt 26){$colL="A"+[char]([int][char]'A'+$col-27)}; $letters=@(); $n=$col; while($n -gt 0){$m=($n-1)%26; $letters=[char]([int][char]'A'+$m)+$letters; $n=[int](($n-1)/26)}; $ref="$($letters -join '')$r"; $line+=("$ref`:$(GetRole $ref)") }; Write-Output ("R{0}: {1}" -f $r, ($line -join ' ')) }

Write-Output ""
Write-Output "Row 17-29 O-R:"
foreach($r in 17..29){ $line=@(); foreach($col in 15..18){ $letters=@(); $n=$col; while($n -gt 0){$m=($n-1)%26; $letters=[char]([int][char]'A'+$m)+$letters; $n=[int](($n-1)/26)}; $ref="$($letters -join '')$r"; $line+=("$ref`:$(GetRole $ref)") }; Write-Output ("R{0}: {1}" -f $r, ($line -join ' ')) }

Write-Output ""
Write-Output "Fill palette from styles.xml (non-none):"
$fills | Select-Object -Skip 1 -Unique | ForEach-Object { if($_ -ne 'none'){ "$(ClassifyFill $_) = $_" } }
