$folder = Get-Location

$files = Get-ChildItem -Path $folder -Filter *.txt -File

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다."
    exit
}

# UTF-16 LE(BOM 포함)
$encoding = [System.Text.Encoding]::Unicode

foreach ($file in $files) {

    try {

        # 파일 내용 읽기
        $content = [System.IO.File]::ReadAllText($file.FullName)

        # UTF-16 LE(BOM 포함)로 저장
        [System.IO.File]::WriteAllText(
            $file.FullName,
            $content,
            $encoding
        )

        Write-Host "변환 완료 : $($file.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "변환 실패 : $($file.Name)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "전체 변환 완료"