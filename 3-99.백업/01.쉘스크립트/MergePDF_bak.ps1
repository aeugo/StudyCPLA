$folder = Get-Location
$output = Join-Path $folder "Merged.pdf"

# Ghostscript 위치
$gs = Get-ChildItem "C:\Program Files\gs\*\bin\gswin64c.exe" |
      Sort-Object FullName -Descending |
      Select-Object -First 1

if(!$gs){
    Write-Host "Ghostscript가 설치되어 있지 않습니다."
    exit
}

$pdfs = Get-ChildItem *.pdf | Sort-Object Name

if($pdfs.Count -eq 0){
    Write-Host "PDF가 없습니다."
    exit
}

$args = @(
"-dBATCH",
"-dNOPAUSE",
"-q",
"-sDEVICE=pdfwrite",
"-sOutputFile=$output"
)

$args += $pdfs.FullName

& $gs.FullName @args

Write-Host "완료 : $output"