# =================================================================
# UTF-16 텍스트 파일 내 "2. 배경" 및 "3. 목적" 섹션 순서/번호 교체 스크립트
# =================================================================

$ErrorActionPreference = "Stop"

# 현재 작업 디렉토리 위치 수집
$folder = Get-Location
$files = @(Get-ChildItem -LiteralPath $folder -Filter *.txt -File)

if ($files.Count -eq 0) {
    Write-Host "처리할 텍스트 파일(*.txt)이 존재하지 않습니다." -ForegroundColor Yellow
    exit
}

$totalCount = $files.Count
$processedCount = 0

Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "    텍스트 파일 섹션 순서 및 번호 변경 프로세스 시작     " -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host ""

foreach ($file in $files) {
    $processedCount++
    
    # UTF-16 (Unicode) 인코딩으로 파일 전체 텍스트 수집
    $rawContent = Get-Content -LiteralPath $file.FullName -Raw -Encoding Unicode

    # [핵심 보정] 각 섹션의 '제목+본문 내용'만 엄격히 수집하는 정규표현식
    # $1: 2. 배경 (제목 ~ 본문 끝)
    # $2: 2. 배경과 3. 목적 사이의 여백/개행
    # $3: 3. 목적 (제목 ~ 본문 끝)
    $fullPattern = "(?ms)(^[ \t]*2\.\s*배경.*?(?=\r?\n[ \t]*3\.\s*목적))(\r?\n[ \t]*)(^[ \t]*3\.\s*목적.*?(?=\r?\n\r?\n[ \t]*[Ⅰ-ⅩIVXLCDM]+\.|\r?\n[ \t]*4\.|\Z))"

    if ($rawContent -match $fullPattern) {
        $bgBlock = $Matches[1]        # 2. 배경 블록
        $middleSpace = $Matches[2]    # 두 섹션 사이의 개행
        $purposeBlock = $Matches[3]   # 3. 목적 블록

        # 목차 번호 교체 (${1} 구문 사용으로 백레퍼런스 오류 방지)
        $newBgBlock = $bgBlock -replace "(?m)^([ \t]*)2\.\s*배경", '${1}3. 배경'
        $newPurposeBlock = $purposeBlock -replace "(?m)^([ \t]*)3\.\s*목적", '${1}2. 목적'

        # 순서만 교체하고 사이 개행 구조는 원본 그대로 유지
        $replacedBlock = $newPurposeBlock + $middleSpace + $newBgBlock

        # 원본 영역을 정확히 1:1 교체
        $updatedContent = $rawContent.Replace($Matches[0], $replacedBlock)

        # UTF-16 (Unicode) 인코딩으로 파일 저장
        Set-Content -LiteralPath $file.FullName -Value $updatedContent -Encoding Unicode -NoNewline

        Write-Host "[$processedCount/$totalCount] $($file.Name) 수정 완료" -ForegroundColor Green
    }
    else {
        Write-Host "[$processedCount/$totalCount] $($file.Name) (변경 대상 양식 미일치 - 건너뜀)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "                  모든 작업이 완료되었습니다.            " -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan