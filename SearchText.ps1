$folder = Get-Location

# 검색할 문자열 입력
$keyword = Read-Host "검색할 문자열을 입력하세요"

if ([string]::IsNullOrWhiteSpace($keyword)) {
    Write-Host "검색 문자열을 입력하지 않았습니다." -ForegroundColor Red
    exit
}

# 모든 하위 폴더의 *.txt 파일까지 탐색
$files = Get-ChildItem -LiteralPath $folder -Filter *.txt -File -Recurse

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다." -ForegroundColor Yellow
    exit
}

$found = $false

# 정규식 패턴 생성 (특수문자 포함 키워드 방어)
$regexPattern = [regex]::Escape($keyword)

foreach ($file in $files) {

    # UTF-8 기준 텍스트 파일 읽기
    $lines = Get-Content -LiteralPath $file.FullName -Encoding UTF-8

    for ($i = 0; $i -lt $lines.Count; $i++) {

        $lineText = $lines[$i]

        if ($lineText.Contains($keyword)) {

            # 상대 경로 계산
            $relativePath = $file.FullName.Replace($folder.Path, "").TrimStart("\", "/")

            Write-Host "파일 : $relativePath" -ForegroundColor Yellow
            Write-Host "줄   : $($i + 1)"
            Write-Host "내용 : " -NoNewline

            # -------------------------------------------------------------
            # [핵심] 검색어를 기준으로 문장을 분할하여 키워드만 색상 강조 출력
            # -------------------------------------------------------------
            $parts = [regex]::Split($lineText, "($regexPattern)", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

            foreach ($part in $parts) {
                if ($part -eq $keyword) {
                    # 검색어 일치 부분: 노란 배경에 빨간 글자(형광펜 효과)로 강조
                    Write-Host $part -ForegroundColor Red -NoNewline #-BackgroundColor Yellow
                }
                else {
                    # 일반 텍스트 부분
                    Write-Host $part -NoNewline
                }
            }
            Write-Host "" # 줄바꿈 처리
            Write-Host ""

            $found = $true
        }
    }
}

if (-not $found) {
    Write-Host "'$keyword' 문자열을 포함한 파일이 없습니다." -ForegroundColor Green
}
else {
    Write-Host "검색이 완료되었습니다." -ForegroundColor Cyan
}