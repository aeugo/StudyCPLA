$folder = Get-Location
$afterFolder = Join-Path $folder "AFTER"

if (!(Test-Path -LiteralPath $afterFolder)) {
    New-Item -ItemType Directory -Path $afterFolder | Out-Null
}

# AFTER 폴더 내부 파일은 탐색 대상에서 완전히 제외
$files = Get-ChildItem -LiteralPath $folder -Filter "*.txt" -File -Recurse | Where-Object { $_.FullName -notlike "$afterFolder*" }

$totalCount = $files.Count

if ($totalCount -eq 0) {
    Write-Host "TXT 파일이 존재하지 않습니다." -ForegroundColor Yellow
    exit
}

Write-Host "총 $totalCount 개 파일에 대해 첫 행 치환 및 AFTER 폴더 저장 작업을 시작합니다..." -ForegroundColor Cyan
Write-Host ""

$currentCount = 0

foreach ($file in $files) {

    $currentCount++

    # 진행과정 표시 : [현재파일 번호/전체파일 수] 현재 진행중인 파일 명
    Write-Host "[$currentCount/$totalCount] $($file.Name)" -ForegroundColor Cyan

    # UTF-16 LE (PowerShell 기준 Unicode) 인코딩으로 파일 읽기
    $lines = Get-Content -LiteralPath $file.FullName -Encoding Unicode
    
    # 빈 파일이거나 내용이 없는 경우 스킵
    if ($null -eq $lines -or $lines.Count -eq 0) { continue }

    # 확장자를 제외한 순수 파일명 추출 및 <파일명.01> 타이틀 생성
    $baseName = $file.BaseName
    $newHeader = "<$baseName.01>"

    # 1. 기존 첫 행을 삭제하고 새 타이틀로 교체하는 라인 재구성
    if ($lines.Count -gt 1) {
        # 2번째 행(인덱스 1)부터 마지막 행까지 추출 후 새 타이틀 결합
        $remainingLines = $lines[1..($lines.Count - 1)]
        $newLines = @($newHeader) + $remainingLines
    } else {
        # 원본 파일이 단 한 줄만 존재하는 경우 첫 행을 바로 교체
        $newLines = @($newHeader)
    }

    # 원본 폴더 기준 상대 경로 계산 및 AFTER 폴더 내 저장 경로 생성
    $relativePath = $file.FullName.Substring($folder.Path.Length).TrimStart([char[]]"\/")
    $targetPath = Join-Path $afterFolder $relativePath
    $targetDir = Split-Path -Path $targetPath -Parent

    # AFTER 폴더 내 하위 디렉터리가 없을 경우 자동 생성
    if (!(Test-Path -LiteralPath $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    # -------------------------------------------------------------
    # [핵심] UTF-16 LE 저장 & 파일 끝 빈행 발생 방지
    # -------------------------------------------------------------
    $fileContent = $newLines -join "`r`n"
    [System.IO.File]::WriteAllText($targetPath, $fileContent, [System.Text.Encoding]::Unicode)
}

Write-Host ""
Write-Host "=============================================================="
Write-Host " 첫 행 치환 및 AFTER 폴더 저장 작업이 완료되었습니다. (UTF-16 LE 저장)" -ForegroundColor Cyan
Write-Host "=============================================================="