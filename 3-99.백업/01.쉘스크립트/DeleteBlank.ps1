$folder = Get-Location

$files = Get-ChildItem -Path $folder -Filter *.txt -File

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다."
    exit
}

$changed = 0

foreach ($file in $files) {

    # 파일 전체 읽기
    $text = [System.IO.File]::ReadAllText(
        $file.FullName,
        [System.Text.Encoding]::UTF8
    )

    $original = $text

    # 파일 끝의 빈 줄(공백 포함) 제거
    while ($text -match "(\r?\n)[`t ]*(\r?\n)?$") {
        $text = $text -replace "(\r?\n)[`t ]*(\r?\n)?$", ""
    }

    if ($text -ne $original) {

        [System.IO.File]::WriteAllText(
            $file.FullName,
            $text,
            [System.Text.Encoding]::UTF8
        )

        Write-Host "수정 : $($file.Name)" -ForegroundColor Yellow
        $changed++
    }
}

Write-Host ""
Write-Host "총 $changed 개 파일 수정 완료."