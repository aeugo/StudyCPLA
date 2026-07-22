$folder = Get-Location

$files = Get-ChildItem -Path $folder -Filter *.txt -File

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다."
    exit
}

# 제1조, 제10조의2, 제20조제1항, 제20조제1항제3호
$pattern = '제\d+조(의\d+)?(제\d+항)?(제\d+호)?'

$found = $false

# ANSI 색상 코드
$Red   = "$([char]27)[31m"
$Reset = "$([char]27)[0m"

foreach ($file in $files) {

    $lines = Get-Content -LiteralPath $file.FullName -Encoding UTF8

    for ($i = 0; $i -lt $lines.Count; $i++) {

        if ($lines[$i] -match $pattern) {

            Write-Host "파일 : $($file.Name)" -ForegroundColor Yellow
            Write-Host "줄   : $($i + 1)"

            # 제○조 부분만 빨간색
            $output = [regex]::Replace(
                $lines[$i],
                $pattern,
                {
                    param($m)
                    "$Red$($m.Value)$Reset"
                }
            )

            Write-Host "내용 : $output"
            Write-Host ""

            $found = $true
        }
    }
}

if (-not $found) {
    Write-Host "법조문이 포함된 파일이 없습니다." -ForegroundColor Green
}
else {
    Write-Host "검색 완료." -ForegroundColor Cyan
}