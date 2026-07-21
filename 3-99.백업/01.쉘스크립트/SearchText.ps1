$folder = Get-Location

# 검색할 문자열 입력
$keyword = Read-Host "검색할 문자열을 입력하세요"

if ([string]::IsNullOrWhiteSpace($keyword)) {
    Write-Host "검색 문자열을 입력하지 않았습니다." -ForegroundColor Red
    exit
}

$files = Get-ChildItem -Path $folder -Filter *.txt -File

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다."
    exit
}

$found = $false

foreach ($file in $files) {

    $lines = Get-Content -LiteralPath $file.FullName -Encoding UTF8

    for ($i = 0; $i -lt $lines.Count; $i++) {

        if ($lines[$i].Contains($keyword)) {

            Write-Host "파일 : $($file.Name)" -ForegroundColor Yellow
            Write-Host "줄   : $($i + 1)"
            Write-Host "내용 : $($lines[$i])"
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