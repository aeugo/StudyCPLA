# 현재 디렉터리의 모든 .txt 파일 탐색
Get-ChildItem -Path . -Filter "*.txt" -File | ForEach-Object {
    $filePath =$_.FullName
    
    # UTF-16 LE(Unicode) 인코딩 지정
    $encoding = [System.Text.Encoding]::Unicode

    # 파일 전체 내용을 인코딩에 맞게 읽기
    $content = [System.IO.File]::ReadAllText($filePath,$encoding)

    if (-not [string]::IsNullOrEmpty($content)) {
        # 줄 바꿈 기호 기준 분리 (기존 줄바꿈 포맷 유지)
        $lines =$content -split "\r?\n"

        if ($lines.Count -gt 0) {
            # 첫 번째 줄의 ".01>" 문자열을 ".00>"로 치환
            $lines[0] =$lines[0] -replace '\.00>', '.01>'
            
            # 원본 파일의 개행 문자열(CRLF)을 유지하면서 하나의 텍스트로 결합
            $newContent =$lines -join "`r`n"

            # 자동 빈 행 추가 없이 UTF-16 LE 인코딩으로 파일 쓰기
            [System.IO.File]::WriteAllText($filePath, $newContent,$encoding)
            Write-Host "처리 완료: $($_.Name)"
        }
    }
}