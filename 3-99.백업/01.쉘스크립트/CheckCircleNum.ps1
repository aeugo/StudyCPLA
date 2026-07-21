$files = Get-ChildItem -Path $folder -Filter *.txt -File

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다."
    exit
}

$found = $false

foreach ($file in $files) {

    $lines = Get-Content -LiteralPath $file.FullName -Encoding UTF8

    for ($i = 0; $i -lt $lines.Count; $i++) {

        $line = $lines[$i]

        # 줄 시작이 아닌 위치에 ①이 있는 경우
        if ($line -match '\S.*①') {

            Write-Host "파일 : $($file.Name)" -ForegroundColor Yellow
            Write-Host "줄   : $($i + 1)"
            Write-Host "내용 : $line"
            Write-Host ""

            $found = $true
        }
    }
}

if (-not $found) {
    Write-Host "문장 중간에 '①'가 포함된 파일이 없습니다." -ForegroundColor Green
}