$inputDir = "C:\Path\To\Your\HTML\Files"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    $htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Use verbatim strings to avoid quote and symbol issues
    $htmlContent = [regex]::Replace($htmlContent, '(?i)\s*style\s*=\s*(".*?"|\'.*?\')', '')
    $htmlContent = [regex]::Replace($htmlContent, '(?i)\s*class\s*=\s*(".*?"|\'.*?\')', '')
    $htmlContent = [regex]::Replace($htmlContent, '(?is)<!--.*?-->', '')
    $htmlContent = [regex]::Replace($htmlContent, '(?i)<(span|div|font)[^>]*>\s*</\1>', '')

    # Remove empty or whitespace-only lines
    $htmlLines = $htmlContent -split "`n" | Where-Object { $_.Trim() -ne "" }
    $htmlContent = $htmlLines -join "`r`n"

    # Save cleaned HTML
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
