# 현재 디렉터리의 모든 .txt 파일 탐색
Get-ChildItem -Filter *.txt | ForEach-Object {
    # 확장자를 제외한 순수 파일명 추출
    $baseName = $_.BaseName
    
    # 1행에 입력할 텍스트 생성 형식: <파일명.01>
    $newContent = "<${baseName}.01>"
    
    # 내용을 덮어쓰고 UTF-16 LE(Unicode) 인코딩으로 저장
    Set-Content -Path $_.FullName -Value $newContent -Encoding Unicode
    
    Write-Host "처리 완료: $($_.Name)"
}

Write-Host "모든 텍스트 파일 처리가 완료되었습니다."