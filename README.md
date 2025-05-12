$inputDir = "C:\path\to\docs"
$outputDir = "C:\path\to\html"
$wdFormatFilteredHTML = 10

if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false

Get-ChildItem -Path $inputDir -Filter *.docx | ForEach-Object {
    $doc = $word.Documents.Open($_.FullName)
    $outputPath = Join-Path $outputDir ($_.BaseName + ".html")
    $doc.SaveAs([ref] $outputPath, [ref] $wdFormatFilteredHTML)
    $doc.Close()
    Write-Host "Exported: $outputPath"
