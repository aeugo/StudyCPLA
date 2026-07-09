# ToPdf.ps1
# NOTE:
# This is a starter generated from our conversation.
# Replace the "Apply-Highlights" function body with your finalized highlighting logic if desired.

$folder = Get-Location
$afterFolder = Join-Path $folder "AFTER"
if (!(Test-Path $afterFolder)) { New-Item -ItemType Directory -Path $afterFolder | Out-Null }

$files = Get-ChildItem -Path $folder -Filter *.txt | Sort-Object Name
$total = $files.Count
if ($total -eq 0) { Write-Host "TXT 파일이 없습니다."; exit }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

function Highlight-Literal {
    param($doc,[string]$text,[int]$color)
    $r=$doc.Content
    $f=$r.Find
    $f.ClearFormatting()
    $f.Text=$text
    $f.MatchWildcards=$false
    $f.Wrap=0
    while($f.Execute()){
        $r.Shading.BackgroundPatternColor=$color
        $r.Collapse(0)
        $r.End=$doc.Content.End
    }
}

function Apply-Highlights {
    param($doc)

    foreach($t in @("Ⅰ.","Ⅱ.","Ⅲ.","Ⅳ.","Ⅴ.","Ⅵ.","Ⅶ.","Ⅷ.","Ⅸ.","Ⅹ.")){
        Highlight-Literal $doc $t 255
    }
    foreach($i in 1..99){
        Highlight-Literal $doc " $i." 49407
        Highlight-Literal $doc "($i)" 5296274
    }
    foreach($c in @("①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩","⑪","⑫","⑬","⑭","⑮","⑯","⑰","⑱","⑲","⑳")){
        Highlight-Literal $doc $c 16767996
    }
}

function Apply-BoldFormatting {
    param($doc)
    #----------------------------------------
    # 전체 Bold
    #----------------------------------------
    $doc.Content.Font.Bold = $true
    #----------------------------------------
    # '－'로 시작하는 문단만 Bold 해제
    #----------------------------------------
    foreach($para in $doc.Paragraphs)
    {
        $text = $para.Range.Text.Trim()

        if($text.StartsWith("－"))
        {
            $para.Range.Font.Bold = $false
        }
    }
}

$idx=1
try{
foreach($file in $files){
 Write-Host "[$idx/$total] 처리중 : $($file.Name)"
Write-Host "1. Word 변환"
 $doc=$word.Documents.Open($file.FullName,$false,$true)
 $range=$doc.Content
 $range.Font.Name="함초롬돋움"
 $range.Font.Size=8
 foreach($sec in $doc.Sections){
   $sec.PageSetup.TopMargin=$word.CentimetersToPoints(1)
   $sec.PageSetup.BottomMargin=$word.CentimetersToPoints(1)
   $sec.PageSetup.LeftMargin=$word.CentimetersToPoints(1.3)
   $sec.PageSetup.RightMargin=$word.CentimetersToPoints(1.3)
 }
 $range.ParagraphFormat.LineSpacingRule=0
 $range.ParagraphFormat.SpaceBefore=0
 $range.ParagraphFormat.SpaceAfter=0

 Write-Host "  1-1. Bold 적용"
Apply-BoldFormatting $doc

Write-Host "  1-2. 목차 컬러 적용"
Apply-Highlights $doc

 $pdf=Join-Path $afterFolder ($file.BaseName+".pdf")
Write-Host "2. PDF 저장"
 $doc.ExportAsFixedFormat($pdf,17)
 $doc.Close($false)
Write-Host "3. 완료"
 $idx++
}
} finally {
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)|Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
}
Write-Host "완료"