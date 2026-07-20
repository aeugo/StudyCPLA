$folder = Get-Location

$files = Get-ChildItem -Path $folder -Filter *.txt -File

$found = $false

foreach ($file in $files) {

    $text = [System.IO.File]::ReadAllText($file.FullName)

    # 마지막 줄이 빈 줄인지 검사 (CRLF, LF 모두 지원)
    if ($text -match "(\r?\n)[ \t]*\z") {

        Write-Host "파일 : $($file.Name)" -ForegroundColor Yellow
        $found = $true
    }
}

if (-not $found) {
    Write-Host "마지막 행이 빈 행인 파일이 없습니다."
}