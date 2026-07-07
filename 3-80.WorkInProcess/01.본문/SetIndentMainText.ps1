# ============================================
# TXT 일괄 변환
# Main + FormatHierarchicalIndentation 통합
# ============================================

$sourceFolder = Get-Location
$outFolder = Join-Path $sourceFolder "AFTER"

if (!(Test-Path $outFolder)) {
    New-Item -ItemType Directory -Path $outFolder | Out-Null
}

$files = Get-ChildItem $sourceFolder -Filter *.txt

$total = $files.Count
$index = 0

foreach($file in $files){

    $index++

    Write-Host "[$index/$total] 처리중 : $($file.Name)"

    $text = [System.IO.File]::ReadAllText(
        $file.FullName,
        [System.Text.Encoding]::UTF8
    )

    #--------------------------------------------------
    # Main()
    #--------------------------------------------------

    $text = $text -replace "`r`n","`n"
    $text = $text -replace "`r","`n"

    while($text.Contains("`n`n")){
        $text = $text.Replace("`n`n","`n")
    }

    $text = $text.Replace("**",'"')

    while($text.Contains('""')){
        $text = $text.Replace('""','"')
    }

    $text = $text.Replace([char]0x318D,[char]0x30FB)
    $text = $text.Replace("·",[char]0x30FB)

    foreach($i in 2..9){
        $text = $text.Replace("제${i}관","`n제${i}관")
    }

    foreach($roman in @("Ⅱ.","Ⅲ.","Ⅳ.","Ⅴ.","Ⅵ.","Ⅶ.","Ⅷ.","Ⅸ.","Ⅹ.")){
        $text = $text.Replace($roman,"`n$roman")
    }

    $text = $text.Replace("`n<","`n`n<")

    foreach($i in 1..10){
        $text = $text.Replace("[$i].","  $i.")
    }

    foreach($i in 1..10){
        $text = $text.Replace("($i)","    ($i)")
    }

    #--------------------------------------------------
    # FormatHierarchicalIndentation()
    #--------------------------------------------------

    $lines = $text -split "`n"

    $result = New-Object System.Collections.Generic.List[string]

    $currentIndent = 0

    foreach($line in $lines){

        $clean = $line.Trim()

        if($clean -eq ""){
            $result.Add("")
            continue
        }

        $indent = 0
        $isTitle = $false

        # <쟁점>, <판결요지>
        if($clean -match '^[<>]'){
            $indent = 0
            $currentIndent = 0
            $isTitle = $true
        }

        # 제#관
        elseif($clean -match '^제\d+관'){
            $indent = 0
            $currentIndent = 0
            $isTitle = $true
        }

        # 로마숫자
        elseif($clean -match '^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]'){
            $indent = 0
            $currentIndent = 0
            $isTitle = $true
        }

        # 숫자
        elseif($clean -match '^\d+\.'){
            $indent = 2
            $currentIndent = 2
            $isTitle = $true
        }

        # (1)
        elseif($clean -match '^\(\d+\)'){
            $indent = 4
            $currentIndent = 4
            $isTitle = $true
        }

        # ①
        elseif($clean -match '^[①②③④⑤⑥⑦⑧⑨⑩]'){
            $indent = 6
            $currentIndent = 6
            $isTitle = $true
        }

        if($isTitle){

            $result.Add((" " * $indent) + $clean)

        }
        else{

            if($clean.StartsWith("－")){
                $clean = $clean.Substring(1).TrimStart()
            }

            $result.Add((" " * ($currentIndent + 1)) + "－" + $clean)

        }

    }

    $output = $result -join "`r`n"

    $savePath = Join-Path $outFolder $file.Name

    [System.IO.File]::WriteAllText(
        $savePath,
        $output,
        [System.Text.Encoding]::UTF8
    )

}

Write-Host ""
Write-Host "==========================="
Write-Host "변환 완료"
Write-Host "==========================="