param(
    [Parameter(Mandatory = $true)]
    [string]$filePath
)

if (!(Test-Path -LiteralPath $filePath)) {
    Write-Error "지정한 파일이 존재하지 않습니다 : $filePath"
    exit
}

# 1. 파일 읽기 및 기본 텍스트 정제
$text = [System.IO.File]::ReadAllText($filePath, [System.Text.Encoding]::Unicode)

# 줄바꿈 단일화 (\r\n, \r -> \n)
$text = $text -replace "`r`n", "`n"
$text = $text -replace "`r", "`n"

# 특수문자 치환 (요청 구문 적용)
$text = $text -replace ""“", "‘"
$text = $text -replace ""”", "’"
$text = $text.Replace([char]0x318D, [char]0x30FB) # U+318D(ㆍ) -> U+30FB(・)

$lines = $text -split "`n"

if ($lines.Count -eq 0) {
    Write-Host "처리할 내용이 없습니다." -ForegroundColor Yellow
    exit
}

# ------------------------------------------------------
# 2. 첫 줄(법령명) 〈 〉 처리
# ------------------------------------------------------
$firstTextIndex = -1
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i].Trim() -ne "") {
        $firstTextIndex = $i
        break
    }
}

if ($firstTextIndex -ne -1) {
    $firstText = $lines[$firstTextIndex].Trim()
    if (-not $firstText.StartsWith("〈")) {
        $lines[$firstTextIndex] = "〈" + $firstText + "〉"
    }
}

# ------------------------------------------------------
# 3. 원본 빈 행 유지 및 법령 체계별 들여쓰기 정제
# ------------------------------------------------------
$result = New-Object System.Collections.Generic.List[string]

# 정규표현식 패턴 정의
$patternPart    = "^제[0-9]+(?:편|장|절|관)(?:\s+.*)?$" # 편, 장, 절, 관

# 조문 제목 비탐욕 매칭 보완 정규표현식
# Matches[1]: 조문 제목 (예: 제854조의2(친생부인의 허가 청구) 또는 제50조[분사무소(分事務所) 설치의 등기])
# Matches[2]: 조문 제목 직후의 본문 전체 (예: ① 어머니 또는... / 법인이 분사무소를...)
$patternArticle = "^(제[0-9]+조(?:의[0-9]+)?(?:\([^)]+\)|\[[^\]]+\])?)\s*(.+)?$"

$patternHang    = "^([①-⑮])\s*(.+)$"
$patternHo      = "^([0-9]+)\.\s*(.+)$"
$patternHo2     = "^([0-9]+의[0-9]+)\.\s*(.+)$"
$patternMok     = "^([가-힣])\.\s*(.+)$"
$patternSubMok  = "^([0-9]+)\)\s*(.+)$"

foreach ($line in $lines) {
    $txt = $line.Trim()

    # 원본 파일의 빈 행인 경우 그대로 유지하여 추가
    if ($txt -eq "") {
        $result.Add("")
        continue
    }

    # [목차] 제○편, 제○장, 제○절, 제○관
    if ($txt -match $patternPart) {
        $result.Add($txt)
        continue
    }

    # [조문] 제○조 (가지번호 및 괄호/대괄호제목)
    if ($txt -match $patternArticle) {
        $titleStr = $Matches[1].Trim()
        $bodyStr  = if ($Matches[2]) { $Matches[2].Trim() } else { "" }

        # 조문 제목 추가
        $result.Add($titleStr)

        # 본문이 함께 붙어있는 경우 다음 줄에 들여쓰기 2칸 적용하여 추가
        if ($bodyStr -ne "") {
            if ($bodyStr -match $patternHang) {
                $bodyStr = $bodyStr -replace $patternHang, '$1 $2'
            }
            $bodyStr = $bodyStr -replace "①", "① " -replace "①  ", "① "
            $result.Add("  " + $bodyStr)
        }
        continue
    }

    # [항] ① ~ ⑮
    if ($txt -match $patternHang) {
        $txt = "  " + ($txt -replace $patternHang, '$1 $2')
        $txt = $txt -replace "①", "① " -replace "①  ", "① "
        $result.Add($txt)
        continue
    }

    # [호의2] 1의2. 2의2. ...
    if ($txt -match $patternHo2) {
        $txt = "    " + ($txt -replace $patternHo2, '$1. $2')
        $result.Add($txt)
        continue
    }

    # [호] 1. 2. ...
    if ($txt -match $patternHo) {
        $txt = "    " + ($txt -replace $patternHo, '$1. $2')
        $result.Add($txt)
        continue
    }

    # [목] 가. 나. ...
    if ($txt -match $patternMok) {
        $txt = "      " + ($txt -replace $patternMok, '$1. $2')
        $result.Add($txt)
        continue
    }

    # [세목] 1) 2) ...
    if ($txt -match $patternSubMok) {
        $txt = "        " + ($txt -replace $patternSubMok, '$1) $2')
        $result.Add($txt)
        continue
    }

    # 기타 일반 본문
    $result.Add($txt)
}

# ------------------------------------------------------
# 4. 파일 저장
# ------------------------------------------------------
$output = $result -join "`r`n"
[System.IO.File]::WriteAllText($filePath, $output, [System.Text.Encoding]::Unicode)