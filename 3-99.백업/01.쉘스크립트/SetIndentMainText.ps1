param(
    [Parameter(Mandatory = $true)]
    [string]$filePath
)

if (!(Test-Path -LiteralPath $filePath)) {
    Write-Error "지정한 파일이 존재하지 않습니다 : $filePath"
    exit
}

$text = [System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::Unicode)

# ==========================================================
# Main() 텍스트 정제
# ==========================================================

$text =$text -replace "`r`n", "`n"
$text = $text -replace "`r", "`n"

while ($text.Contains("`n`n")) {
    $text = $text.Replace("`n`n", "`n")
}

$text =$text.Replace("**", '"')

while ($text.Contains('""')) {
    $text =$text.Replace('""', '"')
}

while ($text.Contains('▶')) {
    $text =$text.Replace('▶', '－')
}

$text =$text.Replace([char]0x318D, [char]0x30FB)
$text =$text.Replace("·", [char]0x30FB)

foreach ($i in 2..9) {$text = $text.Replace("제${i}관", "`n제${i}관")
}

foreach ($roman in @("Ⅱ.", "Ⅲ.", "Ⅳ.", "Ⅴ.", "Ⅵ.", "Ⅶ.", "Ⅷ.", "Ⅸ.", "Ⅹ.")) {
    $text = $text.Replace($roman, "`n$roman")
}

$text =$text.Replace("`n<", "`n`n<")

# [1].~[10]. 목차 기호 치환 (행두 기반 정규표현식)
$text = [regex]::Replace($text, '(?m)^\s*\[(\d+)\]\.', '  $1.')


# ==========================================================
# FormatHierarchicalIndentation() 들여쓰기 정렬
# ==========================================================

$lines = $text -split "`n"
$result = New-Object System.Collections.Generic.List[string]

$currentIndent = 0$previousTitleType = ""

foreach ($line in $lines) {

    $clean =$line.Trim()

    # 빈 행 통과
    if ($clean -eq "") {
        continue
    }

    # ------------------------------------------------------
    # 제목 종류 판별 (행두 패턴 정밀 매칭)
    # ------------------------------------------------------
    $currentTitleType = ""

    if ($clean -match '^<[^>]+>') {$currentTitleType = "Angle"
    }
    elseif ($clean -match '^제\d+관') {
        $currentTitleType = "Article"
    }
    elseif ($clean -match '^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]') {
        if ($clean -match '^Ⅰ') {$currentTitleType = "Roman1"
        }
        else {
            $currentTitleType = "Roman"
        }
    }
    elseif ($clean -match '^\d+\.\s') {$currentTitleType = "Number"
    }
    elseif ($clean -match '^\(\d+\)') {
        $currentTitleType = "Parenthesis"
    }
    elseif ($clean -match '^[①②③④⑤⑥⑦⑧⑨⑩]') {
        $currentTitleType = "CircleNumber"
    }

    # ------------------------------------------------------
    # 들여쓰기 깊이 결정
    # ------------------------------------------------------
    $indent = 0;
    $isTitle =$false;

    if ($clean -match '^<[^>]+>') {
        $indent = 0;
        $currentIndent = 0;
        $isTitle =$true;
    }
    elseif ($clean -match '^제\d+관') {
        $indent = 0;
        $currentIndent = 0;
        $isTitle =$true;
    }
    elseif ($clean -match '^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]') {
        $indent = 0;
        $currentIndent = 0;
        $isTitle =$true;
    }
    elseif ($clean -match '^\d+\.\s') {
        $indent = 2;
        $currentIndent = 2;
        $isTitle =$true;
    }
    elseif ($clean -match '^\(\d+\)') {
        $indent = 4;
        $currentIndent = 4;
        $isTitle =$true;
    }
    elseif ($clean -match '^[①②③④⑤⑥⑦⑧⑨⑩]') {
        $indent = 6;
        $currentIndent = 6;
        $isTitle =$true;
    }
    elseif ($clean -match '^[㉮㉯㉰㉱㉲㉳㉴㉵㉶㉷㉸㉹㉺㉻]') {
        $indent = 8;
        $currentIndent = 8;
        $isTitle =$true;
    }
    elseif ($clean -match '^[㉠㉡㉢㉣㉤㉥㉦㉧㉨㉩㉪㉫㉬㉭]') {
        $indent = 10;
        $currentIndent = 10;
        $isTitle =$true;
    }

    # ------------------------------------------------------
    # 제목인 경우
    # ------------------------------------------------------
    if ($isTitle) {

        # 로마숫자 목차 앞에 빈 행 1개 추가
        if ($currentTitleType -eq "Roman1" -or $currentTitleType -eq "Roman") {
            $result.Add("")
        }

        # Ⅰ ↔ <...> 제목 사이에 빈 행 1개 추가
        if ($currentTitleType -eq "Angle" -and $previousTitleType -eq "Roman1") {
            $result.Add("")
        }

        $result.Add((" " * $indent) +$clean)
    }

    # ------------------------------------------------------
    # 일반 본문인 경우
    # ------------------------------------------------------
    else {

        if ($clean.StartsWith("－")) {
            $clean =$clean.Substring(1).TrimStart()
        }

        $result.Add((" " * ($currentIndent + 1)) + "－" + $clean)
    }

    # ------------------------------------------------------
    # 직전 제목 종류 저장 (본문이 나오면 관계를 끊음)
    # ------------------------------------------------------
    if ($isTitle) {
        $previousTitleType =$currentTitleType;
    }
    else {
        $previousTitleType = "";
    }

}

# ==========================================================
# 파일 저장 (UTF-16 LE 인코딩 유지)
# ==========================================================

$output =$result -join "`r`n"

[System.IO.File]::WriteAllText(
    $filePath,$output,
    [System.Text.Encoding]::Unicode
)