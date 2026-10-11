# 현재 디렉터리 경로 지정
$folder = Get-Location

# 현재 디렉터리의 모든 .txt 파일 탐색
$files = Get-ChildItem -LiteralPath $folder -Filter "*.txt" -File

if ($files.Count -eq 0) {
    Write-Host "가공할 TXT 파일이 존재하지 않습니다." -ForegroundColor Yellow
    exit
}

# UTF-16 LE(Unicode) 인코딩 지정
$encoding = [System.Text.Encoding]::Unicode

foreach ($file in $files) {
    $filePath = $file.FullName
    
    # UTF-16 LE 인코딩으로 파일 전체 내용 읽기
    $content = [System.IO.File]::ReadAllText($filePath, $encoding)

    if (-not [string]::IsNullOrEmpty($content)) {
        # 줄 바꿈 기호 기준 분리 (CRLF 또는 LF 지원)
        $lines = $content -split "\r?\n"

        if ($lines.Count -gt 0) {
            # 첫 번째 행(인덱스 0)만 추출하고 나머지 행은 제거
            $firstLine = $lines[0]

            # 첫 행만 포함된 내용을 UTF-16 LE 인코딩으로 파일에 쓰기 (추가 빈 행 없음)
            [System.IO.File]::WriteAllText($filePath, $firstLine, $encoding)
            Write-Host "처리 완료 (첫 행 유지): $($file.Name)"
        }
    }
}

Write-Host "`n모든 파일의 첫 행만 남기고 나머지 내용 삭제가 완료되었습니다." -ForegroundColor Green