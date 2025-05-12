$inputDir = "C:\Path\To\Your\HTML\Files"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    $htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Define regex patterns as here-strings to avoid quote errors
    $stylePattern = @'(?i)\s*style\s*=\s*(".*?"|'.*?')'@
    $classPattern = @'(?i)\s*class\s*=\s*(".*?"|'.*?')'@
    $commentPattern = @'(?is)<!--.*?-->'@
    $emptyTagPattern = @'(?i)<(span|div|font)[^>]*>\s*</\1>'@

    # Perform regex replacements
    $htmlContent = [regex]::Replace($htmlContent, $stylePattern, '')
    $htmlContent = [regex]::Replace($htmlContent, $classPattern, '')
    $htmlContent = [regex]::Replace($htmlContent, $commentPattern, '')
    $htmlContent = [regex]::Replace($htmlContent, $emptyTagPattern, '')

    # Remove empty or whitespace-only lines
    $htmlLines = $htmlContent -split "`n" | Where-Object { $_.Trim() -ne "" }
    $htmlContent = $htmlLines -join "`r`n"

    # Save cleaned HTML
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
