$folder = Get-Location

$files = Get-ChildItem -Path $folder -Filter *.txt -File

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다."
    exit
}

$changed = 0

foreach ($file in $files) {

    # 파일 읽기
    $content = Get-Content -LiteralPath $file.FullName -Encoding UTF8 -Raw

    # "[예상]"이 있는 경우만 변경
    if ($content.Contains("[예상]")) {

        $content = $content.Replace("[예상]", "[모의]")

        # UTF-8로 저장
        Set-Content -LiteralPath $file.FullName -Value $content -Encoding UTF8

        Write-Host "수정 완료 : $($file.Name)" #-ForegroundColor Green
        $changed++
    }
}

Write-Host ""
Write-Host "====================="
Write-Host ("총 {0}개의 파일을 수정했습니다." -f $changed)
Write-Host "====================="