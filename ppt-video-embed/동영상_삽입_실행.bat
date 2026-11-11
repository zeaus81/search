<# : batch
@echo off
chcp 65001 >nul
set "EMBED_SCRIPT_DIR=%~dp0"
set "EMBED_SCRIPT_FILE=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -STA -Command "iex ([IO.File]::ReadAllText($env:EMBED_SCRIPT_FILE, [Text.Encoding]::UTF8))"
pause
exit /b
#>
# PPT에 "링크"로 연결된 동영상을 PPT 파일 "안에" 넣어(포함) 메일로 보내도 재생되게 만드는 스크립트
# 사용법: 이 파일을 더블클릭 -> PPT 파일 선택 -> 동영상 폴더 선택
# 원본 PPT는 그대로 두고, 같은 폴더에 "원본이름_동영상포함.pptx" 를 새로 만듭니다.

$PptPath = $null
$VideoFolder = $null

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web

$scriptDir = $env:EMBED_SCRIPT_DIR.TrimEnd('\\')

# 1) PPT 파일 선택
if (-not $PptPath) {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = '동영상을 넣을 PPT 파일을 선택하세요'
    $dlg.Filter = 'PowerPoint (*.pptx)|*.pptx'
    $dlg.InitialDirectory = $scriptDir
    if ($dlg.ShowDialog() -ne 'OK') { Write-Host '취소되었습니다.'; exit 1 }
    $PptPath = $dlg.FileName
}

# 2) 동영상 폴더 선택
if (-not $VideoFolder) {
    $fdlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $fdlg.Description = '동영상(.wmv) 파일들이 들어있는 폴더를 선택하세요 (하위 폴더까지 찾습니다)'
    $fdlg.SelectedPath = Split-Path -Parent $PptPath
    if ($fdlg.ShowDialog() -ne 'OK') { Write-Host '취소되었습니다.'; exit 1 }
    $VideoFolder = $fdlg.SelectedPath
}

Write-Host "PPT     : $PptPath"
Write-Host "동영상  : $VideoFolder"
Write-Host '동영상 파일 목록을 읽는 중...'
$allVideos = @(Get-ChildItem -LiteralPath $VideoFolder -Recurse -File -ErrorAction SilentlyContinue)

function Find-Video([string]$name) {
    $hit = $allVideos | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
    if ($hit) { return $hit.FullName }
    return $null
}

$contentTypes = @{
    'wmv' = 'video/x-ms-wmv'; 'mp4' = 'video/mp4'; 'm4v' = 'video/mp4';
    'avi' = 'video/x-msvideo'; 'mpg' = 'video/mpeg'; 'mpeg' = 'video/mpeg';
    'mov' = 'video/quicktime'; 'asf' = 'video/x-ms-asf'
}

$utf8 = New-Object System.Text.UTF8Encoding($false)

function Read-Entry($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), $utf8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}

function Write-Entry($zip, [string]$name, [string]$text) {
    $old = $zip.GetEntry($name)
    if ($old) { $old.Delete() }
    $e = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
    $sw = New-Object System.IO.StreamWriter($e.Open(), $utf8)
    try { $sw.Write($text) } finally { $sw.Close() }
}

# 3) 결과 파일 준비 (원본 복사)
$outPath = Join-Path (Split-Path -Parent $PptPath) ([System.IO.Path]::GetFileNameWithoutExtension($PptPath) + '_동영상포함.pptx')
Copy-Item -LiteralPath $PptPath -Destination $outPath -Force

$zip = [System.IO.Compression.ZipFile]::Open($outPath, [System.IO.Compression.ZipArchiveMode]::Update)
$embedded = @{}   # 원본경로 -> ppt/media 이름
$usedExt = @{}
$missing = New-Object System.Collections.ArrayList
$okCount = 0
$mediaNo = 1
while ($zip.GetEntry("ppt/media/media$mediaNo.wmv") -or $zip.GetEntry("ppt/media/media$mediaNo.mp4")) { $mediaNo++ }

try {
    $relEntries = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/_rels/slide\d+\.xml\.rels$' } | ForEach-Object { $_.FullName })
    foreach ($relName in $relEntries) {
        $rels = Read-Entry $zip $relName
        $slideName = 'ppt/slides/' + ([System.IO.Path]::GetFileName($relName) -replace '\.rels$', '')
        $slide = Read-Entry $zip $slideName
        if (-not $slide) { continue }

        $relMatches = [regex]::Matches($rels, '<Relationship\b[^>]*/>')
        $maxId = 0
        foreach ($m in [regex]::Matches($rels, 'Id="rId(\d+)"')) { $maxId = [Math]::Max($maxId, [int]$m.Groups[1].Value) }
        $changed = $false

        foreach ($m in $relMatches) {
            $tag = $m.Value
            if ($tag -notmatch 'relationships/video"' -or $tag -notmatch 'TargetMode="External"') { continue }
            $id = [regex]::Match($tag, 'Id="([^"]+)"').Groups[1].Value
            $target = [regex]::Match($tag, 'Target="([^"]+)"').Groups[1].Value
            $decoded = [System.Uri]::UnescapeDataString([System.Web.HttpUtility]::HtmlDecode($target))
            $fileName = ($decoded -split '[\\/]')[-1]
            $ext = ([System.IO.Path]::GetExtension($fileName)).TrimStart('.').ToLower()
            $src = Find-Video $fileName
            if (-not $src) {
                [void]$missing.Add("$slideName : $fileName")
                continue
            }

            if ($embedded.ContainsKey($src)) {
                $mediaName = $embedded[$src]
            } else {
                $mediaName = "media$mediaNo.$ext"; $mediaNo++
                Write-Host "  넣는 중: $fileName"
                [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $src, "ppt/media/$mediaName", [System.IO.Compression.CompressionLevel]::NoCompression)
                $embedded[$src] = $mediaName
            }
            $usedExt[$ext] = $true

            # 링크 -> 파일 내부 동영상으로 변경 + PowerPoint 2010 이상용 media 관계 추가
            $maxId++
            $mediaRelId = "rId$maxId"
            $newTag = "<Relationship Id=`"$id`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/video`" Target=`"../media/$mediaName`"/>" +
                      "<Relationship Id=`"$mediaRelId`" Type=`"http://schemas.microsoft.com/office/2007/relationships/media`" Target=`"../media/$mediaName`"/>"
            $rels = $rels.Replace($tag, $newTag)

            $vf = "<a:videoFile r:link=`"$id`"/>"
            $ext14 = "<p:extLst><p:ext uri=`"{DAA4B4D4-6D27-4009-9C2C-2DEAC5AE6A76}`"><p14:media xmlns:p14=`"http://schemas.microsoft.com/office/powerpoint/2010/main`" r:embed=`"$mediaRelId`"/></p:ext></p:extLst>"
            if ($slide.Contains($vf + '</p:nvPr>')) {
                $slide = $slide.Replace($vf + '</p:nvPr>', $vf + $ext14 + '</p:nvPr>')
            }
            $changed = $true
            $okCount++
        }

        if ($changed) {
            Write-Entry $zip $relName $rels
            Write-Entry $zip $slideName $slide
        }
    }

    # 동영상 확장자 등록
    $ct = Read-Entry $zip '[Content_Types].xml'
    foreach ($ext in $usedExt.Keys) {
        if ($ct -notmatch "Extension=`"$ext`"") {
            $type = $contentTypes[$ext]; if (-not $type) { $type = "video/$ext" }
            $ct = $ct.Replace('<Default ', "<Default Extension=`"$ext`" ContentType=`"$type`"/><Default ")
        }
    }
    Write-Entry $zip '[Content_Types].xml' $ct
}
finally {
    $zip.Dispose()
}

Write-Host ''
Write-Host "완료: 동영상 $okCount 개 포함 -> $outPath"
$sizeMB = [Math]::Round((Get-Item -LiteralPath $outPath).Length / 1MB, 1)
Write-Host "파일 크기: $sizeMB MB"
if ($missing.Count -gt 0) {
    Write-Host ''
    Write-Host "찾지 못한 동영상 $($missing.Count) 개 (해당 영상은 여전히 링크 상태입니다):" -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    Write-Host '-> 이 파일들을 동영상 폴더에 넣고 다시 실행하세요.' -ForegroundColor Yellow
}
