$folder = Get-Location

$outputFile = Join-Path $folder "Merged.txt"

if (Test-Path -LiteralPath $outputFile) {
    Remove-Item -LiteralPath $outputFile -Force
}

$files = Get-ChildItem -Path $folder -Filter *.txt -File |
         Sort-Object Name

foreach ($file in $files) {

    Write-Host "추가 : $($file.Name)"

    $text = [System.IO.File]::ReadAllText($file.FullName)

    [System.IO.File]::AppendAllText(
        $outputFile,
        $text + "`r`n`r`n",
        [System.Text.Encoding]::UTF8
    )
}

Write-Host "완료"