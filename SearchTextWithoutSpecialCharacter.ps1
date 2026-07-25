$folder = Get-Location

# 검색할 문자열 입력
$keyword = Read-Host "검색할 문자열을 입력하세요"

if ([string]::IsNullOrWhiteSpace($keyword)) {
    Write-Host "검색 문자열을 입력하지 않았습니다." -ForegroundColor Red
    exit
}

# 모든 하위 폴더의 *.txt 파일 탐색
$files = Get-ChildItem -LiteralPath $folder -Filter *.txt -File -Recurse

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다." -ForegroundColor Yellow
    exit
}

$found = $false
$totalMatchCount = 0      # 총 발견 건수 카운트
$matchedFiles = @{}        # 키워드가 발견된 파일 목록 (중복 제거용)

$regexPattern = [regex]::Escape($keyword)

foreach ($file in $files) {

    # UTF-8 기준 텍스트 파일 읽기
    $lines = Get-Content -LiteralPath $file.FullName -Encoding UTF-8

    for ($i = 0; $i -lt $lines.Count; $i++) {

        $lineText = $lines[$i]

        if ([string]::IsNullOrWhiteSpace($lineText)) { continue }

        # -------------------------------------------------------------
        # [신규 조건] '－'로 시작하는 본문 행만 검색 대상으로 필터링
        # -------------------------------------------------------------
        if (-not $lineText.Trim().StartsWith("－")) {
            continue
        }

        # -------------------------------------------------------------
        # 1. "", 「」, [], () 감싼 영역 임시 치환 (제외 처리)
        # -------------------------------------------------------------
        $quoteMatches = @()
        
        # 큰따옴표, 겹낫표, 대괄호, 소괄호 매칭 정규식
        $excludePattern = '"[^"]*"|「[^」]*」|\[[^\]]*\]|\([^\)]*\)'
        
        # 검사용 문자열 생성 (제외 기호 내부 내용을 임시 치환)
        $maskIndex = 0
        $maskedLine = [regex]::Replace($lineText, $excludePattern, {
            param($m)
            $script:quoteMatches += $m.Value
            $replacement = "___EXCLUDED_QUOTE_${maskIndex}___"
            $script:maskIndex++
            return $replacement
        })

        # -------------------------------------------------------------
        # 2. 치환된 문자열에서 키워드 포함 여부 검사
        # -------------------------------------------------------------
        if ($maskedLine -match $regexPattern) {

            # 집계 카운터 증가
            $totalMatchCount++
            $matchedFiles[$file.FullName] = $true

            # 상대 경로 계산
            $relativePath = $file.FullName.Replace($folder.Path, "").TrimStart("\", "/")

            Write-Host "파일 : $relativePath" -ForegroundColor Yellow
            Write-Host "줄   : $($i + 1)"
            Write-Host "내용 : " -NoNewline

            # 마스킹된 문장에서 키워드 기준으로 분할
            $parts = [regex]::Split($maskedLine, "($regexPattern)", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

            foreach ($part in $parts) {
                # 검색어와 일치하는 경우 (대소문자 구분 없이 비교)
                if ($part -eq $keyword -or $part -ieq $keyword) {
                    Write-Host $part -ForegroundColor Red -NoNewline
                }
                else {
                    # 복원 처리: 마스킹해두었던 원래 문장 복원
                    $restoredPart = $part
                    for ($k = 0; $k -lt $quoteMatches.Count; $k++) {
                        $placeholder = "___EXCLUDED_QUOTE_${k}___"
                        $restoredPart = $restoredPart.Replace($placeholder, $quoteMatches[$k])
                    }
                    Write-Host $restoredPart -NoNewline
                }
            }
            
            Write-Host "" # 줄바꿈
            Write-Host ""

            $found = $true
        }
    }
}

# -------------------------------------------------------------
# 3. 최종 완료 메시지 및 통계 출력
# -------------------------------------------------------------
if (-not $found) {
    Write-Host "'$keyword' 문자열을 포함한 '－' 시작 본문 검색 결과가 없습니다." -ForegroundColor Green
}
else {
    $matchedFileCount = $matchedFiles.Count
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host "검색이 완료되었습니다. (총 $matchedFileCount 개 파일에서 $totalMatchCount 건 발견)" -ForegroundColor Cyan
    Write-Host "==============================================================" -ForegroundColor Cyan
}