$inputDir = "C:\Path\To\Your\HTML\Files"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    $htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Remove inline styles (only double-quoted to avoid parser errors)
    $htmlContent = [regex]::Replace($htmlContent, '(?i)\s*style\s*=\s*\".*?\"', '')

    # Remove class attributes (only double-quoted)
    $htmlContent = [regex]::Replace($htmlContent, '(?i)\s*class\s*=\s*\".*?\"', '')

    # Remove HTML comments
    $htmlContent = [regex]::Replace($htmlContent, '(?s)<!--.*?-->', '')

    # Remove empty span/div/font tags
    $htmlContent = [regex]::Replace($htmlContent, '(?i)<(span|div|font)[^>]*>\s*</\1>', '')

    # Remove empty lines
    $htmlLines = $htmlContent -split "`n" | Where-Object { $_.Trim() -ne "" }
    $htmlContent = $htmlLines -join "`r`n"

    # Save cleaned HTML
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
