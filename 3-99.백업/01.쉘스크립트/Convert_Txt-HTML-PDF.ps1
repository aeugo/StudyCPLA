$folder = Get-Location
$afterFolder = Join-Path $folder "AFTER"
$mergedOutputPath = Join-Path $afterFolder "Merged.pdf"

if (!(Test-Path -LiteralPath $afterFolder)) { 
    New-Item -ItemType Directory -Path $afterFolder | Out-Null 
}

# -------------------------------------------------------------
# [단계 1] 사용자 선택 기능 분기 입력창
# -------------------------------------------------------------
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host "             PDF 변환 및 병합 자동화 프로세스            " -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan

$keepHighlight = ""
while ($keepHighlight -notmatch "^[YN]$") {
    $keepHighlight = (Read-Host "1. 쟁점(〈〉), 법령(""), 판례(“”) 등의 배경 강조를 유지하겠습니까? (Y/N)").ToUpper().Trim()
}

$deleteOriginals = ""
while ($deleteOriginals -notmatch "^[YN]$") {
    $deleteOriginals = (Read-Host "2. PDF 파일 병합 후 원본 PDF 파일을 삭제하겠습니까? (Y/N)").ToUpper().Trim()
}
Write-Host "---------------------------------------------------------"

# -------------------------------------------------------------
# [환경 검증] 가용 인프라(Edge 및 Ghostscript) 확인
# -------------------------------------------------------------
$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (!(Test-Path -LiteralPath $edgePath)) {
    $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
}
if (!(Test-Path -LiteralPath $edgePath)) {
    Write-Host "[오류] Microsoft Edge를 찾을 수 없어 PDF 변환을 진행할 수 없습니다." -ForegroundColor Red
    exit
}

$gs = Get-ChildItem "C:\Program Files\gs\*\bin\gswin64c.exe" | Sort-Object FullName -Descending | Select-Object -First 1
if (!$gs) {
    Write-Host "[오류] Ghostscript가 설치되어 있지 않아 PDF 병합을 진행할 수 없습니다." -ForegroundColor Red
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
# [단계 2] HTML 변환 및 마크업 엔진 함수 (N 옵션 시 특수문자 삭제 보완)
# -------------------------------------------------------------
function Convert-TxtToHtml {
    param(
        [string]$filePath,
        [string]$highlightOption
    )

    $lines = Get-Content -LiteralPath $filePath -Encoding UTF-8
    $htmlLines = @()
    $inDoubleAngle = $false

    foreach ($line in $lines) {
        $cleanLine = $line
        $cleanLine = [System.Net.WebUtility]::HtmlDecode($cleanLine)

        if ($highlightOption -eq "Y") {
            # Y 옵션: 배경 강조 적용 및 《》 내부 보호 조치 진행
            if ($cleanLine.Contains("《")) { $inDoubleAngle = $true }

            if ($inDoubleAngle) {
                $cleanLine = $cleanLine.Replace('"', [string][char]0x0060)
                $cleanLine = $cleanLine.Replace([string][char]0x201C, [string][char]0x0060)
                $cleanLine = $cleanLine.Replace([string][char]0x201D, [string][char]0x0060)
            }

            if ($cleanLine.Contains("》")) { $inDoubleAngle = $false }

            $cleanLine = $cleanLine.Replace("★", "<span class='star-red'>★</span>")
            $cleanLine = $cleanLine.Replace("☆", "<span class='star-red'>☆</span>")
            $cleanLine = [regex]::Replace($cleanLine, '〈.*?〉', { param($m) "<span class='case-blue'>$($m.Value)</span>" })
            $cleanLine = [regex]::Replace($cleanLine, '"([^"]+)"', '<span class="quote-purple">"$1"</span>')
            $cleanLine = [regex]::Replace($cleanLine, '“([^”]+)”', '<span class="quote-pink">“$1”</span>')

            # 기본 작은따옴표만 제거
            $cleanLine = $cleanLine.Replace([string][char]39, "")
            $cleanLine = $cleanLine.Replace([string][char]0x2018, "")
            $cleanLine = $cleanLine.Replace([string][char]0x2019, "")
        } 
        else {
            # [핵심 보완] N 옵션: 배경 강조 배제 및 모든 기존 특수문자(따옴표/꺾쇠류) 일괄 완전 소거
            $cleanLine = $cleanLine.Replace([string][char]34, "")        # 일반 큰따옴표 (")
            $cleanLine = $cleanLine.Replace([string][char]39, "")        # 일반 작은따옴표 (')
            $cleanLine = $cleanLine.Replace([string][char]0x201C, "")    # 유니코드 여는 큰따옴표 (“)
            $cleanLine = $cleanLine.Replace([string][char]0x201D, "")    # 유니코드 닫는 큰따옴표 (”)
            $cleanLine = $cleanLine.Replace([string][char]0x2018, "")    # 유니코드 여는 작은따옴표 (‘)
            $cleanLine = $cleanLine.Replace([string][char]0x2019, "")    # 유니코드 닫는 작은따옴표 (’)
            $cleanLine = $cleanLine.Replace([string][char]0x3008, "")    # 홑화살괄호 여는문자 (〈)
            $cleanLine = $cleanLine.Replace([string][char]0x3009, "")    # 홑화살괄호 닫는문자 (〉)
        }

        $cleanTrimmed = $cleanLine.Trim()

        if ([string]::IsNullOrEmpty($cleanTrimmed)) {
            $htmlLines += "<p class='blank-line'>&nbsp;</p>"
            continue
        }

        # 문단 선두 분류 스타일 적용
        $pClass = "bold-text"
        if ($cleanTrimmed.StartsWith("－")) {
            $pClass = "normal-text"
        }
        if ($cleanTrimmed -match '^<문제\d') {
            $pClass = "problem-title"
        }

        # 목차 기호 치환 처리 (정규식 매칭)
        $processedText = $cleanLine
        $processedText = [regex]::Replace($processedText, '^( *)([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.)', '$1<span class="idx-roman">$2</span>')
        $processedText = [regex]::Replace($processedText, '^( *)(\d{1,2}\.)', '$1<span class="idx-bracket">$2</span>')
        $processedText = [regex]::Replace($processedText, '^( *)(\(\d{1,2}\))', '$1<span class="idx-parenthesis">$2</span>')
        $processedText = [regex]::Replace($processedText, '^( *)([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])', '$1<span class="idx-circle">$2</span>')

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
            .problem-title { color:#0000FF; font-weight:bold; }
            .idx-roman { background-color: #FF0000; padding: 2px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; }
            .idx-bracket { background-color: #FFB400; padding: 2px 4px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; }
            .idx-parenthesis { background-color: #92D050; padding: 2px 2.5px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; }
            .idx-circle { background-color: #8CDBF8; padding: 2px 3px; margin-right: 3px; border-radius: 1px; display: inline-block; line-height: 1; }
            .star-red { color:#FF0000; font-weight:bold; }
            .case-blue { background:#CCE0FF; padding: 0px 1px 2px 1px; border-radius: 1px; }
            .quote-purple { background:#C39BE1; padding: 0px 1px 2px 1px; border-radius: 1px; }
            .quote-pink { background:#FF89FF; padding: 0px 1px 2px 1px; border-radius: 1px; }
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
$createdPdfPaths = @()
$idx = 1

foreach ($file in $files) {
    Write-Host "[$idx/$total] 처리중 : $($file.Name)"
    
    $firstLine = (Get-Content -LiteralPath $file.FullName -Encoding UTF8 -TotalCount 1).Trim()
    $firstLine = $firstLine.Replace('<', '[').Replace('>', ']')
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) { $firstLine = $firstLine.Replace($c, ' ') }
    $firstLine = ($firstLine -replace '\s+', ' ').Trim()

    Write-Host "  1. HTML 변환 및 스타일 가공 (배경 강조: $keepHighlight)" -ForegroundColor Gray
    $htmlResult = Convert-TxtToHtml $file.FullName $keepHighlight
    
    $tempHtmlPath = Join-Path $folder "$($file.BaseName).html"
    [System.IO.File]::WriteAllText($tempHtmlPath, $htmlResult, [System.Text.Encoding]::UTF8)

    Write-Host "  2. 백그라운드 PDF 인쇄 진행" -ForegroundColor Gray
    $pdfName = "$firstLine$($file.BaseName).pdf"
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

    if (Test-Path -LiteralPath $pdfPath) {
        $createdPdfPaths += $pdfPath
    }

    Write-Host "  3. 변환 완료" #-ForegroundColor Green
	Write-Host ""
    $idx++
}

# -------------------------------------------------------------
# [단계 3] Ghostscript 활용 PDF 병합 (Start-Process 분리 안정화 버전)
# -------------------------------------------------------------
Write-Host "`n[단계 3] Ghostscript 결합 및 병합 작업 시작..." #-ForegroundColor Cyan

$mergeList = Get-ChildItem -LiteralPath $afterFolder -Filter *.pdf | 
             Where-Object { $_.FullName -ne $mergedOutputPath } | 
             Sort-Object Name

if ($mergeList.Count -eq 0) {
    Write-Host "병합할 PDF 파일이 존재하지 않아 종료합니다." -ForegroundColor Yellow
    exit
}

$gsPsi = New-Object System.Diagnostics.ProcessStartInfo
$gsPsi.FileName = $gs.FullName
$gsPsi.UseShellExecute = $false
$gsPsi.CreateNoWindow = $true

[void]$gsPsi.ArgumentList.Add("-dBATCH")
[void]$gsPsi.ArgumentList.Add("-dNOPAUSE")
[void]$gsPsi.ArgumentList.Add("-q")
[void]$gsPsi.ArgumentList.Add("-sDEVICE=pdfwrite")
[void]$gsPsi.ArgumentList.Add("-sOutputFile=$mergedOutputPath")

foreach ($pdf in $mergeList) {
    [void]$gsPsi.ArgumentList.Add($pdf.FullName)
}

$gsProcess = New-Object System.Diagnostics.Process
$gsProcess.StartInfo = $gsPsi
$null = $gsProcess.Start()
$gsProcess.WaitForExit()

if (Test-Path -LiteralPath $mergedOutputPath) {
    Write-Host ">> 전체 병합 완료 : $mergedOutputPath" #-ForegroundColor Green
    
    if ($deleteOriginals -eq "Y") {
        Write-Host ">> 2. 옵션(Y)에 따라 개별 원본 PDF 파일 청소를 시작합니다." #-ForegroundColor Yellow
        foreach ($pdf in $mergeList) {
            if (Test-Path -LiteralPath $pdf.FullName) {
                Remove-Item -LiteralPath $pdf.FullName -Force
            }
        }
        Write-Host ">> 개별 원본 PDF 파일 삭제 완료." #-ForegroundColor Green
    } else {
        Write-Host ">> 2. 옵션(N)에 따라 개별 원본 PDF 파일들을 그대로 보존합니다." #-ForegroundColor Gray
    }
} else {
    Write-Host "[오류] PDF 병합 중 예상치 못한 문제가 발생했습니다." #-ForegroundColor Red
}

Write-Host "`n=========================================================" #-ForegroundColor Cyan
Write-Host " 모든 자동화 변환 및 제어 공정이 성공적으로 종결되었습니다." #-ForegroundColor Cyan
Write-Host "=========================================================" #-ForegroundColor Cyan