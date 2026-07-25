# ==================================================
# DOC 일괄 처리
# 1. 모든 DOC 열기
# 2. SetIndentLegalText 실행
# 3. 저장
# 4. TXT 폴더에 UTF-8 TXT 저장
# ==================================================

$folder = Get-Location

# TXT 폴더 생성
$txtFolder = Join-Path $folder "AFTER"

if (!(Test-Path $txtFolder)) {
    New-Item -Path $txtFolder -ItemType Directory | Out-Null
}

# DOC 파일 목록
$files = Get-ChildItem -Path $folder -Filter *.doc
$total = $files.Count
$index = 0

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try
{
    foreach ($file in $files)
    {
        $index++

        Write-Host "[$index/$total] 처리중 : $($file.Name)"

        $doc = $word.Documents.Open($file.FullName)

        # VBA 실행
        Write-Host "[$index/$total] 1. 매크로 실행"
        $word.Run("SetIndentLegalText") #C:\Users\MainUser\AppData\Roaming\Microsoft\Templates\Normal.dotm

        # 저장
        Write-Host "[$index/$total] 2. WORD 저장"
        $doc.Save()

        # TXT 저장 경로
        $txtPath = Join-Path $txtFolder ($file.BaseName + ".txt")

        Write-Host "[$index/$total] 3. TXT 저장"

        $text = $doc.Content.Text

        [System.IO.File]::WriteAllText(
            $txtPath,
            $text,
            [System.Text.Encoding]::UTF8
        )

        Write-Host "[$index/$total] 4. 닫기"
        $doc.Close()

        Write-Host ""
    }
}
finally
{
    $word.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "========================"
Write-Host "      작업 완료"
Write-Host "========================"