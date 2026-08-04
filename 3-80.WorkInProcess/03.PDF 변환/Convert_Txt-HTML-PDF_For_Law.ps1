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

        $cleanTrimmed = $cleanLine.Trim()

        if ([string]::IsNullOrEmpty($cleanTrimmed)) {
            $htmlLines += '<p class="blank-line">&nbsp;</p>'
            continue
        }

        # 문단 선두 분류 스타일 적용
        $pClass = "bold-text"
        if ($cleanTrimmed.StartsWith("－") -or $line.StartsWith(" ") -or $line.StartsWith("`t") -or $cleanTrimmed.StartsWith("[")) {
            $pClass = "normal-text"
        }

        $processedText = $cleanLine

        # -----------------------------------------------------
        # [Step 1] < > 괄호 텍스트 우선 마크업 보호 (HTML 태그 충돌 방지)
        # -----------------------------------------------------
        $processedText = [regex]::Replace($processedText, '<([^>]+)>', '__BLUE_START__$1__BLUE_END__')

        # -----------------------------------------------------
        # [Step 2] 일반 법령/인용 구문 배경색 치환
        # -----------------------------------------------------
        # 1) 제목 마크업 매핑 (연파랑 강조)
        $processedText = [regex]::Replace($processedText, '〈.*?〉', { param($m) "<span class=`"case-blue`">$($m.Value)</span>" })

        # 1-1) 법령 전체 인용 마크업 매핑 (보라색 강조)
        $processedText = [regex]::Replace($processedText, '「.*?」', { param($m) "<span class=`"corner-bracket`">$($m.Value)</span>" })

        # 2) 지정된 법령/부령 명칭 배경강조 적용
        $processedText = [regex]::Replace($processedText, '(대통령령|총리령|고용노동부령|보건복지부령|행정안전부령)', '<span class="quote-cyan">$1</span>')

        # 3) 문단 선두 목차 요소 치환 (장, 절에도 가지항목 패턴 적용)
        $processedText = [regex]::Replace($processedText, '^( *)(제\d+편)', '$1<span class="idx-part">$2</span>') #편
        $processedText = [regex]::Replace($processedText, '^( *)(제\d+장(?:의\d+)*)', '$1<span class="idx-chapter">$2</span>') #장(가지항목 포함)
        $processedText = [regex]::Replace($processedText, '^( *)(제\d+절(?:의\d+)*)', '$1<span class="idx-section">$2</span>') #절(가지항목 포함)
        $processedText = [regex]::Replace($processedText, '^( *)(제\d+관)', '$1<span class="idx-subsection">$2</span>') #관
        $processedText = [regex]::Replace($processedText, '^( *)(제\d+조(?:의\d+)*)', '$1<span class="idx-article">$2</span>') #조(가지항목 포함)
        $processedText = [regex]::Replace($processedText, '^( *)([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])', '$1<span class="idx-paragraph">$2</span>') #항(원숫자)
        $processedText = [regex]::Replace($processedText, '^( *)(\d{1,3}(?:의\d+)*\.)', '$1<span class="idx-subparagraph">$2</span>') #호(문단선두)
        $processedText = [regex]::Replace($processedText, '^( *)([가-하]\.)', '$1<span class="idx-item">$2</span>') #목(문단선두)
        $processedText = [regex]::Replace($processedText, '^( *)(\d{1,2}\))', '$1<span class="idx-item2">$2</span>') #목하위

        # 4) 본문 내 참조 구문 통합 하이라이트 치환
        #$processedText = [regex]::Replace($processedText, '(?<!<[^>]*)(?!다목적댐|특정단어2)(?:(?:법|영)\s+)?((?:제\d+조(?:의\d+)*|제\d+항|제\d+호(?:의\d+)*)\s+(?:본문|단서|전단|후단)|제\d+조(?:의\d+)*|제\d+항|제\d+호(?:의\d+)*|(?:가|나|다|라|마|바|사|아|자|차|카|타|파|하)목|별지\s*제\d+호(?:의\d+)*\s*서식|별표\s*\d+(?:의\d+)*|제\d+급)', '<span class="idx-article">$0</span>')
        $processedText = [regex]::Replace($processedText, '(?<!<[^>]*)(?<!class="[^"]*)(?!다목적댐|특정단어2)(?:(?:법|영)\s+)?((?:제\d+조(?:의\d+)*|제\d+항|제\d+호(?:의\d+)*)\s+(?:본문|단서|전단|후단)|제\d+편|제\d+장(?:의\\d+)*|제\d+절(?:의\\d+)*|제\d+관|제\d+조(?:의\d+)*|제\d+항|제\d+호(?:의\d+)*|(?:가|나|다|라|마|바|사|아|자|차|카|타|파|하)목|별지\s*제\d+호(?:의\d+)*\s*서식|별표\s*\d+(?:의\d+)*|제\d+급)(?![^<]*<\/span>)', {
            param($m)
            # 이미 태그에 감싸진 상태인 경우 치환 제외
            if ($m.Value -match '<[^>]+>') { return $m.Value }
            return "<span class=`"idx-article`">$($m.Value)</span>"
        })
        
        # -----------------------------------------------------
        # [Step 3] 대괄호 [ ] 영역 매칭 및 내부 모든 HTML 태그 완벽 제거
        # -----------------------------------------------------
        $processedText = [regex]::Replace($processedText, '\[([^\]]+)\]', {
            param($match)
            $inside = $match.Groups[1].Value
            $cleanInside = [regex]::Replace($inside, '<[^>]+>', '')
            return "<span class=`"no-bg`">[$cleanInside]</span>"
        })

        # -----------------------------------------------------
        # [Step 4] 보호된 < > 괄호 텍스트 파란색 글자 최종 복원
        # -----------------------------------------------------
        $processedText = $processedText.Replace('__BLUE_START__', '<span class="quote-blue">&lt;')
        $processedText = $processedText.Replace('__BLUE_END__', '&gt;</span>')

        $htmlLines += "<p class=`"$pClass`">$processedText</p>"
    }

    # HTML 구조 정의
    $htmlContent = @"
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page { size: A4; margin: 1cm 1.3cm 1cm 1.3cm; }
            body { font-family: "함초롬돋움", sans-serif; font-size: 7.7pt; line-height: 1.6; margin: 0; padding: 0; }
            p { margin: 0; padding: 0; white-space: pre-wrap; word-break: break-all; text-align: justify; text-justify: inter-character; }
            .blank-line { height: 8pt; }
            .bold-text { font-weight: bold; }
            .normal-text { font-weight: normal; }

            span {
                vertical-align: baseline;
                line-height: inherit;
            }

            .no-bg, .no-bg * {
                background: none !important;
                background-color: transparent !important;
                background-image: none !important;
                color: inherit;
                display: inline;
            }

            .quote-blue {
                color: #0000FF;
                display: inline;
            }

            .case-blue {
                background: linear-gradient(to top, #CCE0FF 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 0.5px;
                padding-right: 0.5px;
            }
            .quote-cyan {
                background: linear-gradient(to top, #B6FFFF 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 0.5px;
                padding-right: 0.5px;
            }
            /*편*/
            .idx-part {
                background: linear-gradient(to top, #66CCFF 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 0.5px;
                padding-right: 0.5px;
            }
            /*장*/
            .idx-chapter {
                background: linear-gradient(to top, #FF9FA1 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 0.5px;
                padding-right: 0.5px;
            }
            /*절*/
            .idx-section {
                background: linear-gradient(to top, #ADE57B 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 0.5px;
                padding-right: 0.5px;
            }
            /*관*/
            .idx-subsection {
                background: linear-gradient(to top, #DCB406 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 0.5px;
                padding-right: 0.5px;
            }
            /*조*/
            .idx-article {
                background: linear-gradient(to top, #D9E1F2 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 0.5px;
                padding-right: 0.5px;
            }
            /*항*/
            .idx-paragraph { 
                background: linear-gradient(to top, #FCE4D6 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 1.5px;
                padding-right: 2.5px;
            }
            /*호*/
            .idx-subparagraph {
                background: linear-gradient(to top, #EDEDED 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 3px;
                padding-right: 2.5px;
            }
            /*목*/
            .idx-item {
                background: linear-gradient(to top, #DDEBF7 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 0.2px;
                padding-right: 0.2px;
            }
            /*반괄호*/
            .idx-item2 {
                background: linear-gradient(to top, #E2EFDA 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 3px;
                padding-right: 2.3px;
            }
            /*법령 전체 인용*/
            .corner-bracket { 
                background: linear-gradient(to top, #C382D2 90%, transparent 90%);
                display: inline;
                box-decoration-break: clone;
                -webkit-box-decoration-break: clone;
                padding-top: 1px;
                padding-bottom: 1.5px;
                padding-left: 0.5px;
                padding-right: 0.5px;
            } 
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

    #Start-Sleep -Milliseconds 300

    if (Test-Path -LiteralPath $tempHtmlPath) { Remove-Item -LiteralPath $tempHtmlPath -Force }

    Write-Host "  3. 변환 완료"
    Write-Host ""
    $idx++
}

Write-Host "=============================================================="
Write-Host " 모든 PDF 변환 공정이 성공적으로 종결되었습니다."
Write-Host "=============================================================="
Write-Host ""