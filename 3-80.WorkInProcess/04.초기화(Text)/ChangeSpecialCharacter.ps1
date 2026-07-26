$folder = Get-Location

# 현재 폴더 및 하위 폴더의 모든 *.txt 파일 탐색
$files = Get-ChildItem -LiteralPath $folder -Filter *.txt -File -Recurse

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 존재하지 않습니다." -ForegroundColor Yellow
    exit
}

Write-Host "총 $($files.Count)개 파일에 대해 홑따옴표(' ') 치환 작업을 시작합니다... (인코딩: UTF-16 LE)" -ForegroundColor Cyan
Write-Host ""

$processedFileCount = 0
$totalReplaceCount = 0

foreach ($file in $files) {

    # UTF-16 LE (PowerShell 기준 Unicode) 인코딩으로 파일 읽기
    $lines = Get-Content -LiteralPath $file.FullName -Encoding Unicode
    $newLines = @()
    $fileModified = $false

    foreach ($line in $lines) {

        if ([string]::IsNullOrWhiteSpace($line)) {
            $newLines += $line
            continue
        }

        # -------------------------------------------------------------
        # 1. '...' 내부 문자열 임시 마스킹 (삭제/치환 금지 구역 보호)
        # -------------------------------------------------------------
        $protectedMatches = @()
        
        # '...' (홑따옴표로 감싸진 영역) 정규식
        $singleQuoteBlockPattern = "'[^']*'"
        
        $maskIndex = 0
        $maskedLine = [regex]::Replace($line, $singleQuoteBlockPattern, {
            param($m)
            $script:protectedMatches += $m.Value
            $replacement = "___PROTECTED_SINGLE_QUOTE_${maskIndex}___"
            $script:maskIndex++
            return $replacement
        })

        # -------------------------------------------------------------
        # 2. 보호 구역 외의 ‘’ 기호를 모두 ' 로 변경
        # -------------------------------------------------------------
        $convertedLine = $maskedLine.Replace("‘", "'").Replace("’", "'")

        # -------------------------------------------------------------
        # 3. 마스킹해두었던 '...' 구문 원래대로 복원
        # -------------------------------------------------------------
        for ($k = 0; $k -lt $protectedMatches.Count; $k++) {
            $placeholder = "___PROTECTED_SINGLE_QUOTE_${k}___"
            $convertedLine = $convertedLine.Replace($placeholder, $protectedMatches[$k])
        }

        # 변경 사항 여부 체크
        if ($line -ne $convertedLine) {
            $fileModified = $true
            $totalReplaceCount++
        }

        $newLines += $convertedLine
    }

    # -------------------------------------------------------------
    # [핵심 보완] WriteAllText와 -join 조합을 사용하여 마지막 줄 끝 빈 행 방지
    # -------------------------------------------------------------
    if ($fileModified) {
        $fileContent = $newLines -join "`r`n"
        [System.IO.File]::WriteAllText($file.FullName, $fileContent, [System.Text.Encoding]::Unicode)
        
        $relativePath = $file.FullName.Replace($folder.Path, "").TrimStart("\", "/")
        Write-Host "[수정완료] $relativePath" -ForegroundColor Green
        $processedFileCount++
    }
}

Write-Host ""
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " 치환 작업이 종결되었습니다. (UTF-16 LE 저장)" -ForegroundColor Cyan
Write-Host " (총 $processedFileCount 개 파일 수정 완료 / $totalReplaceCount 개 행 치환)" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan