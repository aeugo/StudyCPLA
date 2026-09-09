param(
    [string]$filePath
)

if (!(Test-Path $filePath)) { exit }

$text = [System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::Unicode)

# Main() 텍스트 정제
$text = $text -replace "`r`n", "`n"
$text = $text -replace "`r", "`n"

while ($text.Contains("`n`n")) {
    $text = $text.Replace("`n`n", "`n")
}

$text = $text.Replace("**", '"')

while ($text.Contains('""')) {
    $text = $text.Replace('""', '"')
}

while ($text.Contains('▶')) {
    $text = $text.Replace('▶', '－')
}

$text = $text.Replace([char]0x318D, [char]0x30FB)
$text = $text.Replace("·", [char]0x30FB)

foreach ($i in 2..9) {
    $text = $text.Replace("제${i}관", "`n제${i}관")
}

foreach ($roman in @("Ⅱ.", "Ⅲ.", "Ⅳ.", "Ⅴ.", "Ⅵ.", "Ⅶ.", "Ⅷ.", "Ⅸ.", "Ⅹ.")) {
    $text = $text.Replace($roman, "`n$roman")
}

$text = $text.Replace("`n<", "`n`n<")

foreach ($i in 1..10) {
    $text = $text.Replace("$i.", "  $i.")
}

foreach ($i in 1..10) {
    $text = $text.Replace("($i)", "    ($i)")
}

# FormatHierarchicalIndentation() 들여쓰기 정렬
$lines = $text -split "`n"
$result = New-Object System.Collections.Generic.List[string]
$currentIndent = 0

foreach ($line in $lines) {
    $clean = $line.Trim()

    if ($clean -eq "") {
        $result.Add("")
        continue
    }

    $indent = 0
    $isTitle = $false

    if ($clean -match '^[<>]') {
        $indent = 0; $currentIndent = 0; $isTitle = $true
    }
    elseif ($clean -match '^제\d+관') {
        $indent = 0; $currentIndent = 0; $isTitle = $true
    }
    elseif ($clean -match '^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]') {
        $indent = 0; $currentIndent = 0; $isTitle = $true
    }
    elseif ($clean -match '^\d+\.') {
        $indent = 2; $currentIndent = 2; $isTitle = $true
    }
    elseif ($clean -match '^\(\d+\)') {
        $indent = 4; $currentIndent = 4; $isTitle = $true
    }
    elseif ($clean -match '^[①②③④⑤⑥⑦⑧⑨⑩]') {
        $indent = 6; $currentIndent = 6; $isTitle = $true
    }

    if ($isTitle) {
        $result.Add((" " * $indent) + $clean)
    }
    else {
        if ($clean.StartsWith("－")) {
            $clean = $clean.Substring(1).TrimStart()
        }
        $result.Add((" " * ($currentIndent + 1)) + "－" + $clean)
    }
}

$output = $result -join "`r`n"
[System.IO.File]::WriteAllText($filePath, $output, [System.Text.Encoding]::Unicode)