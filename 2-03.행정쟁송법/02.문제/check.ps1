$folder = Get-Location

$files = Get-ChildItem -Filter *.txt -File

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다."
    exit
}

$found = $false

foreach ($file in $files) {

    # 파일 전체를 하나의 문자열로 읽기
    $content = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8

    # 《...》 안에 ①~⑩이 있는지 검사
    if ($content -match '《[\s\S]*?[①②③④⑤⑥⑦⑧⑨⑩][\s\S]*?》') {

        Write-Host "발견 : $($file.Name)" -ForegroundColor Yellow
        $found = $true
    }
}

if (-not $found) {
    Write-Host "《》 내부에 ①~⑩이 포함된 파일이 없습니다." -ForegroundColor Green
}