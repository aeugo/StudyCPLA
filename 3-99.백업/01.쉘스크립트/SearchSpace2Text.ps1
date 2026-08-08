$folder = Get-Location

# 모든 하위 폴더의 *.txt 파일까지 탐색
$files = Get-ChildItem -LiteralPath $folder -Filter *.txt -File -Recurse

if ($files.Count -eq 0) {
    Write-Host "TXT 파일이 없습니다." -ForegroundColor Yellow
    exit
}

$found = $false

# 2개 이상의 연속된 스페이스바 공백을 찾는 정규식 패턴
$regexPattern = " {2,}"

foreach ($file in $files) {

    # UTF-8 기준 텍스트 파일 읽기
    $lines = Get-Content -LiteralPath $file.FullName -Encoding UTF-8

    for ($i = 0; $i -lt $lines.Count; $i++) {

        $lineText = $lines[$i]

        # 정규표현식 그룹 캡처 : 라인 시작 부분의 들여쓰기 공백($1)과 본문 텍스트($2) 분리
        if ($lineText -match "^(\s*)(.*)$") {

            $leadingSpace = $Matches[1]
            $restOfLine   = $Matches[2]

            # 본문 텍스트($restOfLine) 내부에서만 2개 이상의 연속 공백이 존재하는지 검사
            if ([System.Text.RegularExpressions.Regex]::IsMatch($restOfLine, $regexPattern)) {

                # 상대 경로 계산
                $relativePath = $file.FullName.Replace($folder.Path, "").TrimStart("\", "/")

                Write-Host "파일 : $relativePath" -ForegroundColor Yellow
                Write-Host "줄   : $($i + 1)"
                Write-Host "내용 : " -NoNewline

                # 1. 들여쓰기(선두 공백)는 강조 없이 원본 그대로 출력
                Write-Host $leadingSpace -NoNewline

                # 2. 본문 텍스트 영역만 연속 공백을 기준으로 분할하여 강조 출력
                $parts = [regex]::Split($restOfLine, "($regexPattern)")

                foreach ($part in $parts) {
                    if ([System.Text.RegularExpressions.Regex]::IsMatch($part, "^ {2,}$")) {
                        # 문장 중간의 2개 이상 연속 공백 : 노란 배경에 빨간 글자로 강조
                        Write-Host $part -ForegroundColor Red -BackgroundColor Yellow -NoNewline
                    }
                    else {
                        # 일반 본문 텍스트 부분
                        Write-Host $part -NoNewline
                    }
                }
                Write-Host "" # 줄바꿈 처리
                Write-Host ""

                $found = $true
            }
        }
    }
}

if (-not $found) {
    Write-Host "문장 중간에 빈칸(스페이스바)이 두 개 이상인 텍스트 파일이 없습니다." -ForegroundColor Green
}
else {
    Write-Host "문장 중간 연속 공백 검사가 완료되었습니다." -ForegroundColor Cyan
}