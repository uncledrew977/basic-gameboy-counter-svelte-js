$inputDir = "C:\Path\To\Input\DOCX Files"
$outputDir = "C:\Path\To\Output\HTML"
$converterScript = "C:\Path\To\convert-single.ps1"

# Make sure output directory exists
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

Get-ChildItem -Path $inputDir -Filter *.docx | ForEach-Object {
    $inputFile = $_.FullName
    $outputFile = Join-Path $outputDir ($_.BaseName + ".html")

    $cmd = "& `"$converterScript`" -InputPath `"$inputFile`" -OutputPath `"$outputFile`""
    Start-Process powershell -ArgumentList "-NoProfile", "-ExecutionPolicy Bypass", "-Command $cmd" -Wait -NoNewWindow
}
