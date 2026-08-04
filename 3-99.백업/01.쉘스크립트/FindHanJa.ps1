# param(
#     [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
#     [string]$FolderPath,
    
#     [string]$FileFilter = "*.txt"
# )

# # 1. 지정 폴더 존재 여부 확인
# if (!(Test-Path -LiteralPath $FolderPath)) {
#     Write-Host "[오류] 지정한 폴더 경로를 찾을 수 없습니다: $FolderPath" -ForegroundColor Red
#     exit
# }


$folder = Get-Location
$files = Get-ChildItem -Path $folder -Filter *.txt -File

# 2. 한자 탐색용 유니코드 정규표현식 패턴
$hanjaRegex = '[\u4E00-\u9FFF\u3400-\u4DBF\uF900-\uFAFF]'

# 3. 폴더 내 파일 목록 수집 (하위 폴더 포함)
$files = Get-ChildItem -Path $FolderPath -Filter $FileFilter -Recurse -File

if ($files.Count -eq 0) {
    Write-Host "[알림] '$FolderPath' 폴더 내에 '$FileFilter' 조건에 맞는 파일이 없습니다." -ForegroundColor Yellow
    exit
}

$totalHanjaFiles = 0

Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " 폴더 내 전체 텍스트 파일 한자 탐색 시작" -ForegroundColor Cyan
Write-Host " 대상 폴더: $FolderPath" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan

# 4. 파일별 탐색 루프
foreach ($file in $files) {
    $lines = Get-Content -LiteralPath $file.FullName -Encoding UTF-8 -ErrorAction SilentlyContinue
    $lineNumber = 1
    $fileHanjaCount = 0
    $fileOutputHeaderPrinted = $false

    foreach ($line in $lines) {
        if ($line -match $hanjaRegex) {
            if (!$fileOutputHeaderPrinted) {
                Write-Host "`n[파일] $($file.FullName)" -ForegroundColor Green
                $fileOutputHeaderPrinted = $true
                $totalHanjaFiles++
            }
            Write-Host "  [Line $lineNumber] $line" -ForegroundColor Yellow
            $fileHanjaCount++
        }
        $lineNumber++
    }
}

Write-Host "`n==============================================================" -ForegroundColor Cyan
if ($totalHanjaFiles -eq 0) {
    Write-Host "폴더 내 어떤 파일에서도 한자가 감지되지 않았습니다." -ForegroundColor Green
} else {
    Write-Host "총 $totalHanjaFiles 개의 파일에서 한자가 포함된 행이 감지되었습니다." -ForegroundColor Green
}
Write-Host "==============================================================" -ForegroundColor Cyan