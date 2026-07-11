# ============================================
# <제 로 시작하지 않는 TXT 파일 찾기
# ============================================

$folder = Get-Location

$files = Get-ChildItem -Path $folder -Filter *.txt | Sort-Object Name

$total = $files.Count
$index = 1

foreach($file in $files)
{
    Write-Host "[$index/$total] 검사중 : $($file.Name)"

    # UTF-8 읽기
    $lines = Get-Content -LiteralPath $file.FullName -Encoding UTF8

    # 첫 번째 빈 줄이 아닌 줄 찾기
    $firstLine = $null

    foreach($line in $lines)
    {
        $text = $line.Trim()

        if($text -ne "")
        {
            $firstLine = $text
            break
        }
    }

    if($null -eq $firstLine)
    {
        Write-Host "  빈 파일 : $($file.Name)" -ForegroundColor Yellow
    }
    elseif(-not $firstLine.StartsWith("<제"))
    {
        Write-Host "  [발견] $($file.Name)" -ForegroundColor Red
        Write-Host "         첫 줄 : $firstLine"
    }

    $index++
}

Write-Host ""
Write-Host "=========================="
Write-Host "검사 완료"
Write-Host "=========================="