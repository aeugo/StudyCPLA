$folder = Get-Location
$afterFolder = Join-Path $folder "AFTER"
if (!(Test-Path $afterFolder)) { 
    New-Item -ItemType Directory -Path $afterFolder | Out-Null 
}

$files = Get-ChildItem -Path $folder -Filter *.txt | Sort-Object Name
$total = $files.Count
if ($total -eq 0) {
    Write-Host "TXT 파일이 없습니다."
    exit 
}

# Edge 브라우저 경로 확인 (64비트 및 32비트 경로 체크)
$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if (!(Test-Path $edgePath)) {
    $edgePath = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
}

if (!(Test-Path $edgePath)) {
    Write-Host "Microsoft Edge를 찾을 수 없어 PDF 변환을 진행할 수 없습니다."
    exit
}

function Convert-TxtToHtml {
    param([string]$filePath)

    # 텍스트 파일 읽기 (UTF-8 기준)
    $lines = Get-Content -LiteralPath $filePath -Encoding UTF-8
    $htmlLines = @()

    foreach ($line in $lines) {
		$cleanLine = $line

		# HTML Entity 복원
		$cleanLine = [System.Net.WebUtility]::HtmlDecode($cleanLine)

		# 특수문자 제거
		$cleanLine = $cleanLine.Replace([string][char]34, "")        # 일반 큰따옴표 (")
        $cleanLine = $cleanLine.Replace([string][char]39, "")        # 일반 작은따옴표 (')
        $cleanLine = $cleanLine.Replace([string][char]0x201C, "")    # 유니코드 여는 큰따옴표 (“)
        $cleanLine = $cleanLine.Replace([string][char]0x201D, "")    # 유니코드 닫는 큰따옴표 (”)
        $cleanLine = $cleanLine.Replace([string][char]0x2018, "")    # 유니코드 여는 작은따옴표 (‘)
        $cleanLine = $cleanLine.Replace([string][char]0x2019, "")    # 유니코드 닫는 작은따옴표 (’)
        $cleanLine = $cleanLine.Replace([string][char]0x3008, "")    # 홑화살괄호 여는문자 (〈)
        $cleanLine = $cleanLine.Replace([string][char]0x3009, "")    # 홑화살괄호 닫는문자 (〉)

		# ★, ☆ 빨간색 표시
		$cleanLine = $cleanLine.Replace(
			"★",
			"<span class='star-red'>★</span>"
		)

		$cleanLine = $cleanLine.Replace(
			"☆",
			"<span class='star-red'>☆</span>"
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
		text-align: justify;       /* 우측면 정렬을 균일하게 맞추는 양끝 정렬 속성 */
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
	.problem-title{
		color:#0000FF;
		font-weight:bold;
	}
    
    /* 목차 배경 및 세부 여백 조절 디자인 */
    .idx-roman {
        background-color: #FF0000;
        /*color: #FFFFFF;*/
        padding: 2px 2px 2px 2px;
        margin-right: 3px;
        border-radius: 1px;
        display: inline-block;
        line-height: 1;
    }
    .idx-bracket {
        background-color: #FFB400;
        /*color: #FFFFFF;*/
        padding: 2px 4px 2px 4px;
        margin-right: 3px;
        border-radius: 1px;
        display: inline-block;
        line-height: 1;
    }
    .idx-parenthesis {
        background-color: #92D050;
        /*color: #FFFFFF;*/
        padding: 2px 2.5px 2px 2.5px;
        margin-right: 3px;
        border-radius: 1px;
        display: inline-block;
        line-height: 1;
    }
    .idx-circle {
        background-color: #8CDBF8;
        /*color: #FFFFFF;*/
        padding: 2px 3px 2px 3px;
        margin-right: 3px;
        border-radius: 1px;
        display: inline-block;
        line-height: 1;
    }
	/*별표 강조*/
	.star-red{
		color:#FF0000;
		font-weight:bold;
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
    
    # 1. HTML 문자열 생성
    Write-Host "  1. HTML 변환 및 스타일 적용"
    $htmlResult = Convert-TxtToHtml $file.FullName
    
    # 임시 HTML 파일 저장 경로
    $tempHtmlPath = Join-Path $folder "$($file.BaseName)_temp.html"
    [System.IO.File]::WriteAllText($tempHtmlPath, $htmlResult, [System.Text.Encoding]::UTF8)

    # 2. Edge Headless 기능을 이용한 PDF 저장
    Write-Host "  2. PDF 변환 진행"
    $pdfPath = Join-Path $afterFolder ($file.BaseName + ".pdf")
    
    # Edge 가용 인자 설정 (--no-pdf-header-footer 필수 설정으로 상하단 파일정보 출력 제거)
    $arguments = @(
        "--headless",
        "--disable-gpu",
        "--no-pdf-header-footer",
        "--print-to-pdf=`"$pdfPath`"",
        "`"$tempHtmlPath`""
    )
    
    # 프로세스 실행 및 종료 대기
	$psi = New-Object System.Diagnostics.ProcessStartInfo
	$psi.FileName = $edgePath
	$psi.Arguments = ($arguments -join " ")
	$psi.UseShellExecute = $false
	$psi.CreateNoWindow = $true
	$psi.RedirectStandardError = $true
	$psi.RedirectStandardOutput = $true

	$process = New-Object System.Diagnostics.Process
	$process.StartInfo = $psi

	$null = $process.Start()

	$stderr = $process.StandardError.ReadToEnd()
	$stdout = $process.StandardOutput.ReadToEnd()

	$process.WaitForExit()

	if ($stderr) {
		$stderr | Set-Content "$env:TEMP\edge_err.log"
	}

	# Edge 브라우저가 파일 핸들을 완전히 해제하도록 단시간 대기 조치
    Start-Sleep -Milliseconds 300

    # 리터럴 경로 설정을 활용하여 대괄호가 포함된 파일의 무조건적인 삭제 강제
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