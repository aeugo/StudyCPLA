$folder = Get-Location
$afterFolder = Join-Path $folder "AFTER"

if (!(Test-Path -LiteralPath $afterFolder)) { 
    New-Item -ItemType Directory -Path $afterFolder | Out-Null 
}

# -------------------------------------------------------------
# [환경 검증] 가용 인프라(Edge 브라우저) 확인
# -------------------------------------------------------------
$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (!(Test-Path -LiteralPath $edgePath)) {
    $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
}
if (!(Test-Path -LiteralPath $edgePath)) {
    Write-Host "[오류] Microsoft Edge를 찾을 수 없어 PDF 변환을 진행할 수 없습니다." -ForegroundColor Red
    exit
}

# 변환 대상 TXT 파일 탐색
$files = Get-ChildItem -LiteralPath $folder -Filter *.txt | Sort-Object Name
$total = $files.Count
if ($total -eq 0) {
    Write-Host "가공할 TXT 파일이 존재하지 않습니다." -ForegroundColor Yellow
    exit 
}

# -------------------------------------------------------------
# [HTML 변환 및 마크업 엔진 함수]
# -------------------------------------------------------------
function Convert-TxtToHtml {
    param(
        [string]$filePath
    )

    $lines = Get-Content -LiteralPath $filePath -Encoding UTF-8
    $htmlLines = @()

    foreach ($line in $lines) {
        $cleanLine = $line
        $cleanLine = [System.Net.WebUtility]::HtmlDecode($cleanLine)

        # 강조 마크업 매핑 (법령제목 등)
        $cleanLine = [regex]::Replace($cleanLine, '「.*?」', { param($m) "<span class='quote-purple'>$($m.Value)</span>" })

        $cleanTrimmed = $cleanLine.Trim()

        if ([string]::IsNullOrEmpty($cleanTrimmed)) {
            $htmlLines += "<p class='blank-line'>&nbsp;</p>"
            continue
        }

        # 문단 선두 분류 스타일 적용 (첫 칸이 공백이거나 '－'로 시작하면 normal-text)
        $pClass = "bold-text"
        if ($cleanTrimmed.StartsWith("－") -or $line.StartsWith(" ") -or $line.StartsWith("`t")) {
            $pClass = "normal-text"
        }

        # 목차 및 편 기호 치환 처리 (정규식 매칭)
        $processedText = $cleanLine
        $processedText = [regex]::Replace($processedText, '^( *)(제\d+편)', '$1<span class="idx-part">$2</span>') #편
		$processedText = [regex]::Replace($processedText, '^( *)(제\d+장)', '$1<span class="idx-chapter">$2</span>') #장
		$processedText = [regex]::Replace($processedText, '^( *)(제\d+절)', '$1<span class="idx-section">$2</span>') #절
		$processedText = [regex]::Replace($processedText, '^( *)(제\d+관)', '$1<span class="idx-subsection">$2</span>') #관
		$processedText = [regex]::Replace($processedText, '^( *)(제\d+조)', '$1<span class="idx-article">$2</span>') #조
		$processedText = [regex]::Replace($processedText, '^( *)([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])', '$1<span class="idx-paragraph">$2</span>') #항
        $processedText = [regex]::Replace($processedText, '^( *)(\d{1,2}\.)', '$1<span class="idx-subparagraph">$2</span>') #호
		$processedText = [regex]::Replace($processedText, '^( *)([가-하]\.)', '$1<span class="idx-item">$2</span>') #목

        $htmlLines += "<p class='$pClass'>$processedText</p>"
    }

    # HTML 구조 정의 (CSS 렌더링 스펙 포함)
    $htmlContent = @"
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page { size: A4; margin: 1cm 1.3cm 1cm 1.3cm; }
            body { font-family: "함초롬돋움", sans-serif; font-size: 8pt; line-height: 1.6; margin: 0; padding: 0; }
            p { margin: 0; padding: 0; white-space: pre-wrap; word-break: break-all; text-align: justify; text-justify: inter-character; }
            .blank-line { height: 8pt; }
            .bold-text { font-weight: bold; }
            .normal-text { font-weight: normal; }

            .idx-part { background-color: #FF9FA1; padding: 2px 4px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; } /*편*/
            .idx-chapter { background-color: #FF0000; padding: 2px 4px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; } /*장*/
			.idx-section { background-color: #92D050; padding: 2px 4px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; } /*절*/
			.idx-subsection { background-color: #BF8F00; padding: 2px 4px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; } /*관*/
			.idx-article { background-color: #D9E1F2; padding: 2px 4px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; } /*조*/
            .idx-paragraph { background-color: #FCE4D6; padding: 2px 2px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; } /*항*/
            .idx-subparagraph { background-color: #EDEDED; padding: 2px 2px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; } /*호*/
			.idx-item { background-color: #DDEBF7; padding: 2px 2px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; } /*목*/

            .quote-purple { background:#C39BE1; padding: 0px 1px 2px 1px; border-radius: 1px; } /*법령제목*/
        </style>
    </head>
    <body>$($htmlLines -join "`n")</body>
</html>
"@
    return $htmlContent
}

# -------------------------------------------------------------
# [루프 구동] 개별 파일 PDF 변환 수행
# -------------------------------------------------------------
$idx = 1

foreach ($file in $files) {
    Write-Host "[$idx/$total] 처리중 : $($file.Name)"

    Write-Host "  1. HTML 변환 및 스타일 가공" -ForegroundColor Gray
    $htmlResult = Convert-TxtToHtml $file.FullName
    
    $tempHtmlPath = Join-Path $folder "$($file.BaseName).html"
    [System.IO.File]::WriteAllText($tempHtmlPath, $htmlResult, [System.Text.Encoding]::UTF8)

    Write-Host "  2. 백그라운드 PDF 인쇄 진행" -ForegroundColor Gray
    $pdfName = "$($file.BaseName).pdf"
    $pdfPath = Join-Path $afterFolder $pdfName

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $edgePath
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true

    [void]$psi.ArgumentList.Add("--headless")
    [void]$psi.ArgumentList.Add("--disable-gpu")
    [void]$psi.ArgumentList.Add("--no-pdf-header-footer")
    [void]$psi.ArgumentList.Add("--print-to-pdf=$pdfPath")
    [void]$psi.ArgumentList.Add($tempHtmlPath)

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    $null = $process.Start()
    $process.WaitForExit()

    Start-Sleep -Milliseconds 300

    if (Test-Path -LiteralPath $tempHtmlPath) { Remove-Item -LiteralPath $tempHtmlPath -Force }

    Write-Host "  3. 변환 완료"
    Write-Host ""
    $idx++
}

Write-Host "=============================================================="
Write-Host " 모든 PDF 변환 공정이 성공적으로 종결되었습니다."
Write-Host "=============================================================="