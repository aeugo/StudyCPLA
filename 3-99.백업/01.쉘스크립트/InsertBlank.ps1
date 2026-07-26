$folder = Get-Location

# 현재 폴더 및 하위 폴더의 모든 *.txt 파일 탐색
$files = Get-ChildItem -LiteralPath $folder -Filter *.txt -File -Recurse

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 존재하지 않습니다." -ForegroundColor Yellow
    exit
}

Write-Host "총 $($files.Count)개 파일에 대해 제목-로마숫자 간 빈행 검사 작업을 시작합니다..." -ForegroundColor Cyan
Write-Host ""

$processedFileCount = 0
$totalAddedLines = 0

foreach ($file in $files) {

    # UTF-16 LE (PowerShell 기준 Unicode) 인코딩으로 파일 읽기
    $lines = Get-Content -LiteralPath $file.FullName -Encoding Unicode
    
    if ($lines.Count -eq 0) { continue }

    $newLines = @()
    $fileModified = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $currentLine = $lines[$i]
        $newLines += $currentLine

        # 다음 줄이 존재하는지 확인
        if ($i + 1 -lt $lines.Count) {
            $nextLine = $lines[$i + 1]

            # 1. 현재 줄이 '<'로 시작하는 제목 행인지 판단 (앞뒤 공백 제거 기준)
            $isTitleLine = $currentLine.Trim().StartsWith("<")
            
            # 2. 다음 줄이 'Ⅰ' (로마숫자 1)로 시작하는지 판단
            $isNextRomanOne = $nextLine.Trim().StartsWith("Ⅰ")

            # < 로 시작하는 제목 바로 다음 줄이 Ⅰ 로 시작하면 빈행 추가
            if ($isTitleLine -and $isNextRomanOne) {
                $newLines += ""  # 빈행 삽입
                $fileModified = $true
                $totalAddedLines++
            }
        }
    }

    # -------------------------------------------------------------
    # [핵심] UTF-16 LE 저장 & 파일 끝 빈행 발생 방지
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
Write-Host " 빈행 추가 작업이 종결되었습니다. (UTF-16 LE 저장)" -ForegroundColor Cyan
Write-Host " (총 $processedFileCount 개 파일 수정 완료 / $totalAddedLines 개 빈행 추가)" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan