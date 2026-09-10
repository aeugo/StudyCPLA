# param(
#     [string]$filePath
# )

# if (!(Test-Path $filePath)) { exit }

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

foreach ($roman in @(
    "Ⅱ.", "Ⅲ.", "Ⅳ.", "Ⅴ.",
    "Ⅵ.", "Ⅶ.", "Ⅷ.", "Ⅸ.", "Ⅹ."
)) {
    $text = $text.Replace($roman, "`n$roman")
}

$text = $text.Replace("`n<", "`n`n<")

foreach ($i in 1..10) {
    $text = $text.Replace("$i.", "  $i.")
}

foreach ($i in 1..10) {
    $text = $text.Replace("($i)", "    ($i)")
}


# ==========================================================
# FormatHierarchicalIndentation() 들여쓰기 정렬
# ==========================================================

$lines = $text -split "`n"

$result = New-Object System.Collections.Generic.List[string]

$currentIndent = 0

# 직전의 제목 종류 저장
# Roman1  = Ⅰ
# Roman   = Ⅱ~Ⅹ
# Angle   = <...>
# Other   = 기타 제목
$previousTitleType = ""

foreach ($line in $lines) {

    $clean = $line.Trim()

    # ------------------------------------------------------
    # 빈 행
    # ------------------------------------------------------
    if ($clean -eq "") {
        continue
    }


    # ------------------------------------------------------
    # 제목 종류 판별
    # ------------------------------------------------------

    $currentTitleType = ""

    # <...> 제목
    if ($clean -match '^[<>]') {
        $currentTitleType = "Angle"
    }

    # 제n관
    elseif ($clean -match '^제\d+관') {
        $currentTitleType = "Article"
    }

    # 로마숫자
    elseif ($clean -match '^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]') {

        # Ⅰ인지 확인
        if ($clean -match '^Ⅰ') {
            $currentTitleType = "Roman1"
        }
        else {
            $currentTitleType = "Roman"
        }
    }

    # 숫자 제목
    elseif ($clean -match '^\d+\.') {
        $currentTitleType = "Number"
    }

    # (숫자) 제목
    elseif ($clean -match '^\(\d+\)') {
        $currentTitleType = "Parenthesis"
    }

    # ①②③...
    elseif ($clean -match '^[①②③④⑤⑥⑦⑧⑨⑩]') {
        $currentTitleType = "CircleNumber"
    }


    # ------------------------------------------------------
    # 들여쓰기
    # ------------------------------------------------------

    $indent = 0
    $isTitle = $false

    if ($clean -match '^[<>]') {
        $indent = 0
        $currentIndent = 0
        $isTitle = $true
    }

    elseif ($clean -match '^제\d+관') {
        $indent = 0
        $currentIndent = 0
        $isTitle = $true
    }

    elseif ($clean -match '^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]') {
        $indent = 0
        $currentIndent = 0
        $isTitle = $true
    }

    elseif ($clean -match '^\d+\.') {
        $indent = 2
        $currentIndent = 2
        $isTitle = $true
    }

    elseif ($clean -match '^\(\d+\)') {
        $indent = 4
        $currentIndent = 4
        $isTitle = $true
    }

    elseif ($clean -match '^[①②③④⑤⑥⑦⑧⑨⑩]') {
        $indent = 6
        $currentIndent = 6
        $isTitle = $true
    }


    # ------------------------------------------------------
    # 제목인 경우
    # ------------------------------------------------------

    if ($isTitle) {

        # ==================================================
        # 1. 로마숫자 목차 앞에 빈 행 1개 무조건 추가
        # ==================================================

        if ($currentTitleType -eq "Roman1" -or $currentTitleType -eq "Roman") {
            $result.Add("")
        }


        # ==================================================
        # 2. Ⅰ ↔ <...> 제목 사이에 빈 행 1개
        # ==================================================

        if (
            $currentTitleType -eq "Angle" -and
            $previousTitleType -eq "Roman1"
        ) {
            $result.Add("")
        }


        $result.Add((" " * $indent) + $clean)

    }

    # ------------------------------------------------------
    # 일반 본문
    # ------------------------------------------------------

    else {

        if ($clean.StartsWith("－")) {
            $clean = $clean.Substring(1).TrimStart()
        }

        $result.Add((" " * ($currentIndent + 1)) + "－" + $clean)
    }


    # ------------------------------------------------------
    # 직전 제목 종류 저장
    #
    # 본문이 나오면 제목 관계를 끊음
    # ------------------------------------------------------

    if ($isTitle) {
        $previousTitleType = $currentTitleType
    }
    else {
        $previousTitleType = ""
    }

}

# ==========================================================
# 파일 저장
# ==========================================================

$output = $result -join "`r`n"

[System.IO.File]::WriteAllText(
    $filePath,
    $output,
    [System.Text.Encoding]::Unicode
)