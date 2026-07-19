$folder = Get-Location
$afterFolder = Join-Path $folder "AFTER"
if (!(Test-Path -LiteralPath $afterFolder)) { 
    New-Item -ItemType Directory -Path $afterFolder | Out-Null 
}

# 대괄호 포함 파일 탐색을 위해 LiteralPath 규칙 적용
$files = Get-ChildItem -LiteralPath $folder -Filter *.txt | Sort-Object Name
$total = $files.Count
if ($total -eq 0) {
    Write-Host "TXT 파일이 없습니다."
    exit 
}

# Edge 브라우저 경로 확인
$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (!(Test-Path -LiteralPath $edgePath)) {
    $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
}

if (!(Test-Path -LiteralPath $edgePath)) {
    Write-Host "Microsoft Edge를 찾을 수 없어 PDF 변환을 진행할 수 없습니다."
    exit
}

function Convert-TxtToHtml {
    param([string]$filePath)

    # 텍스트 파일 읽기 (UTF-8 기준)
    $lines = Get-Content -LiteralPath $filePath -Encoding UTF-8
    $htmlLines = @()

    # [핵심 보완] 《 와 》 사이 영역을 추적하는 상태 관리 플래그 변수 선언
    $inDoubleAngle = $false

    foreach ($line in $lines) {
        $cleanLine = $line

        # HTML Entity 복원
        $cleanLine = [System.Net.WebUtility]::HtmlDecode($cleanLine)

        # -------------------------------------------------------------
        # [핵심 알고리즘 적용] 《 와 》 범위 안의 따옴표 백틱 처리 영역
        # -------------------------------------------------------------
        if ($cleanLine.Contains("《")) {
            $inDoubleAngle = $true
        }

        if ($inDoubleAngle) {
            # 《 와 》 영역 내부라면 모든 형태의 큰따옴표를 백틱(`)으로 완전 변경
            $cleanLine = $cleanLine.Replace('"', [string][char]0x0060)
            $cleanLine = $cleanLine.Replace([string][char]0x201C, [string][char]0x0060) # 유니코드 “
            $cleanLine = $cleanLine.Replace([string][char]0x201D, [string][char]0x0060) # 유니코드 ”
        }

        if ($cleanLine.Contains("》")) {
            $inDoubleAngle = $false # 닫는 괄호가 발견되면 플래그 비활성화
        }
        # -------------------------------------------------------------

        # 특수문자 제거 (이전 줄바꿈이나 괄호 처리가 꼬이지 않도록 대상 문자만 안전 교체)
        $cleanLine = $cleanLine.Replace([string][char]39, "")        # 일반 작은따옴표 (')
        $cleanLine = $cleanLine.Replace([string][char]0x2018, "")    # 유니코드 여는 작은따옴표 (‘)
        $cleanLine = $cleanLine.Replace([string][char]0x2019, "")    # 유니코드 닫는 작은따옴표 (’)

        #$cleanLine = $cleanLine.Replace([string][char]34, "")        # 일반 큰따옴표 (")
        #$cleanLine = $cleanLine.Replace([string][char]0x201C, "")    # 유니코드 여는 큰따옴표 (“)
        #$cleanLine = $cleanLine.Replace([string][char]0x201D, "")    # 유니코드 닫는 큰따옴표 (”)
        #$cleanLine = $cleanLine.Replace([string][char]0x3008, "")    # 홑화살괄호 여는문자 (〈)
        #$cleanLine = $cleanLine.Replace([string][char]0x3009, "")    # 홑화살괄호 닫는문자 (〉)

        # ★, ☆ 빨간색 표시
        $cleanLine = $cleanLine.Replace("★", "<span class='star-red'>★</span>")
        $cleanLine = $cleanLine.Replace("☆", "<span class='star-red'>☆</span>")

        # 〈...〉 : 쟁점 배경
        $cleanLine = [regex]::Replace(
            $cleanLine,
            '〈.*?〉',
            {
                param($m)
                "<span class='case-blue'>$($m.Value)</span>"
            }
        )

        # "..." : 법령 배경 (치환된 백틱 영역이 마킹되지 않도록 순수한 따옴표만 매칭)
        $cleanLine = [regex]::Replace(
            $cleanLine,
            '"([^"]+)"',
            '<span class="quote-purple">"$1"</span>'
        )

        # “...” : 판례 배경
        $cleanLine = [regex]::Replace(
            $cleanLine,
            '“([^”]+)”',
            '<span class="quote-pink">“$1”</span>'
        )

        $cleanTrimmed = $cleanLine.Trim()

        # 빈 줄 처리
        if ([string]::IsNullOrEmpty($cleanTrimmed)) {
            $htmlLines += "<p class='blank-line'>&nbsp;</p>"
            continue
        }

        # － 로 시작하는 문단은 일반 스타일, 그 외는 볼드 스타일 적용을 위한 클래스 분기
        $pClass = "bold-text"
        if ($cleanTrimmed.StartsWith("－")) {
            $pClass = "normal-text"
        }

        # 문제 제목
        if($cleanTrimmed -match '^<문제\d') {
            $pClass = "problem-title"
        }

        # 목차 기호 치환 처리 (정규식 및 반복문 활용)
        $processedText = $cleanLine

        # 1. 로마숫자 치환 (Ⅰ. ~ Ⅹ.)
        $processedText = [regex]::Replace(
            $processedText,
            '^( *)([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.)',
            '$1<span class="idx-roman">$2</span>'
        )

        # 2. 일반 숫자 치환 (1. ~ 99.)
        $processedText = [regex]::Replace(
            $processedText,
            '^( *)(\d{1,2}\.)',
            '$1<span class="idx-bracket">$2</span>'
        )

        # 3. 괄호 숫자 치환 ((1) ~ (99))
        $processedText = [regex]::Replace(
            $processedText,
            '^( *)(\(\d{1,2}\))',
            '$1<span class="idx-parenthesis">$2</span>'
        )

        # 4. 원숫자 치환 (① ~ ⑳)
        $processedText = [regex]::Replace(
            $processedText,
            '^( *)([①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳])',
            '$1<span class="idx-circle">$2</span>'
        )

        # HTML 문단 태그로 감싸기
        $htmlLines += "<p class='$pClass'>$processedText</p>"
    }

    # 완성된 HTML 구조 생성 (CSS 포함)
    $htmlContent = @"
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page {
                size: A4;
                margin-top: 1cm;
                margin-bottom: 1cm;
                margin-left: 1.3cm;
                margin-right: 1.3cm;
            }
            body {
                font-family: "함초롬돋움", sans-serif;
                font-size: 8pt;
                line-height: 1.6;
                margin: 0;
                padding: 0;
            }
            p {
                margin-top: 0px;
                margin-bottom: 0px;
                padding: 0;
                white-space: pre-wrap;
                word-break: break-all;
                text-align: justify; /* 우측면 정렬을 균일하게 맞추는 양끝 정렬 속성 */
                text-justify: inter-character; /* 글자 간격을 미세 조정하여 공백 불균형 해소 */
            }
            .blank-line {
                height: 8pt;
            }
            .bold-text {
                font-weight: bold;
            }
            .normal-text {
                font-weight: normal;
            }
            .problem-title {
                color:#0000FF;
                font-weight:bold;
            }
            
            /* 목차 배경 및 세부 여백 조절 디자인 */
            .idx-roman {
                background-color: #FF0000;
                padding: 2px 2px 2px 2px;
                margin-right: 3px;
                border-radius: 1px;
                display: inline-block;
                line-height: 1;
            }
            .idx-bracket {
                background-color: #FFB400;
                padding: 2px 4px 2px 4px;
                margin-right: 3px;
                border-radius: 1px;
                display: inline-block;
                line-height: 1;
            }
            .idx-parenthesis {
                background-color: #92D050;
                padding: 2px 2.5px 2px 2.5px;
                margin-right: 3px;
                border-radius: 1px;
                display: inline-block;
                line-height: 1;
            }
            .idx-circle {
                background-color: #8CDBF8;
                padding: 2px 3px 2px 3px;
                margin-right: 3px;
                border-radius: 1px;
                display: inline-block;
                line-height: 1;
            }
            /*별표 강조*/
            .star-red {
                color:#FF0000;
                font-weight:bold;
            }
            /*쟁점 강조*/
            .case-blue {
                background:#CCE0FF;
                padding: 0px 1px 2px 1px;
                border-radius: 1px;
            }
            /* 법령 배경 */
            .quote-purple {
                background:#C39BE1;
                padding: 0px 1px 2px 1px;
                border-radius: 1px;
            }
            /* 판례 배경 */
            .quote-pink {
                background:#FF89FF;
                padding: 0px 1px 2px 1px;
                border-radius: 1px;
            }
        </style>
    </head>
    <body>
        $($htmlLines -join "`n")
    </body>
</html>
"@

    return $htmlContent
}

$idx = 1
foreach ($file in $files) 
{
    Write-Host "[$idx/$total] 처리중 : $($file.Name)"
    
    # 첫 번째 줄 읽기
    $firstLine = (Get-Content -LiteralPath $file.FullName -Encoding UTF8 -TotalCount 1).Trim()

    # < → [ , > → ] 치환
    $firstLine = $firstLine.Replace('<', '[')
    $firstLine = $firstLine.Replace('>', ']')

    # 파일명에 사용할 수 없는 문자 제거
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    foreach($c in $invalidChars){
        $firstLine = $firstLine.Replace($c,' ')
    }

    # 연속 공백 정리
    $firstLine = ($firstLine -replace '\s+',' ').Trim()

    # HTML 생성
    Write-Host "  1. HTML 변환 및 스타일 적용"
    $htmlResult = Convert-TxtToHtml $file.FullName
    
    # 임시 HTML 파일 저장 경로
    $tempHtmlPath = Join-Path $folder "$($file.BaseName).html"
    [System.IO.File]::WriteAllText($tempHtmlPath, $htmlResult, [System.Text.Encoding]::UTF8)

    # 2. Edge Headless 기능을 이용한 PDF 저장
    Write-Host "  2. PDF 변환 진행"
    $pdfName = "$firstLine$($file.BaseName).pdf"
    $pdfPath = Join-Path $afterFolder $pdfName
    
    # [백그라운드 고정] 팝업 창 차단을 위한 Windows GUI 전면 은닉 설정
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $edgePath
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true

    # 안전한 무인(Headless) 다원 배열 인젝션 구조 실행
    [void]$psi.ArgumentList.Add("--headless")
    [void]$psi.ArgumentList.Add("--disable-gpu")
    [void]$psi.ArgumentList.Add("--no-pdf-header-footer")
    [void]$psi.ArgumentList.Add("--print-to-pdf=$pdfPath")
    [void]$psi.ArgumentList.Add($tempHtmlPath)

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi

    $null = $process.Start()

    $stderr = $process.StandardError.ReadToEnd()
    $stdout = $process.StandardOutput.ReadToEnd()

    $process.WaitForExit()

    if ($stderr) {
        $stderr | Set-Content "$env:TEMP\edge_err.log"
    }

    # 파일 안전 점유 해제를 위한 미세 대기 지연
    Start-Sleep -Milliseconds 300

    # 리터럴 경로 설정을 활용한 임시 리포트 완전 소거
    if (Test-Path -LiteralPath $tempHtmlPath) {
        Remove-Item -LiteralPath $tempHtmlPath -Force
    }

    Write-Host "  3. 완료"
    Write-Host ""
    $idx++
}

Write-Host "====================="
Write-Host "전체 완료"
Write-Host "====================="