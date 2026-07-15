$files = Get-ChildItem -Path $folder -Filter *.txt -File

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다."
    exit
}

$found = $false

foreach ($file in $files) {

    $lines = Get-Content -LiteralPath $file.FullName -Encoding UTF8

    for($i = 0; $i -lt $lines.Count; $i++) {

        # 공백 없이 바로 －로 시작하는 행
        if($lines[$i] -match '^－') {

            Write-Host "파일 : $($file.Name)" -ForegroundColor Yellow
            Write-Host "줄   : $($i+1)"
            Write-Host "내용 : $($lines[$i])"
            Write-Host ""

            $found = $true
        }
    }
}

if(-not $found){
    Write-Host "공백 없이 '－'로 시작하는 행이 없습니다." -ForegroundColor Green
}