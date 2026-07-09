# ============================================
# TXT → PDF 일괄 변환
# - 현재 폴더의 TXT 대상
# - AFTER 폴더에 PDF 저장
# - 글꼴 : 맑은 고딕
# - 크기 : 9pt
# ============================================
$folder = Get-Location
$afterFolder = Join-Path $folder "AFTER"

if (!(Test-Path $afterFolder))
{
    New-Item -ItemType Directory -Path $afterFolder | Out-Null
}

$files = Get-ChildItem -Path $folder -Filter *.txt | Sort-Object Name

$total = $files.Count

if ($total -eq 0)
{
    Write-Host "TXT 파일이 없습니다."
    exit
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$index = 1

try
{
    foreach ($file in $files)
    {
        Write-Host ""
        Write-Host "[$index/$total] 처리중 : $($file.Name)"
				Write-Host "1. Word 변환"
        #-----------------------------
        # TXT 열기
        #-----------------------------
        $doc = $word.Documents.Open(
            $file.FullName,
            $false,
            $true
        )

        #-----------------------------
        # 전체 선택
        #-----------------------------
        $range = $doc.Content

        # 글꼴
        $range.Font.Name = "함초롬돋움"
        $range.Font.Size = 8

				foreach($sec in $doc.Sections)
				{
						$sec.PageSetup.TopMargin    = $word.CentimetersToPoints(1)
						$sec.PageSetup.BottomMargin = $word.CentimetersToPoints(1)
						$sec.PageSetup.LeftMargin   = $word.CentimetersToPoints(1.3)
						$sec.PageSetup.RightMargin  = $word.CentimetersToPoints(1.3)
				}

        # 줄간격
        $range.ParagraphFormat.LineSpacingRule = 0    # wdLineSpaceSingle
        $range.ParagraphFormat.SpaceBefore = 0
        $range.ParagraphFormat.SpaceAfter = 0

        # PDF 경로
        $pdfPath = Join-Path $afterFolder ($file.BaseName + ".pdf")

        Write-Host "2. PDF 저장"

        #-----------------------------
        # PDF 저장
        #-----------------------------
        $doc.ExportAsFixedFormat(
            $pdfPath,
            17      # wdExportFormatPDF
        )

        $doc.Close($false)

        Write-Host "3. 완료"

        $index++
    }
}
finally
{
    $word.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "===================================="
Write-Host "모든 PDF 변환 완료"
Write-Host "===================================="