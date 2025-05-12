$inputDir = "C:\Test\Output"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    $htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Remove class attributes (double-quoted)
    $htmlContent = [regex]::Replace($htmlContent, '\sclass\s*=\s*"[^"]*"', '')

    # Remove style attributes (double-quoted)
    $htmlContent = [regex]::Replace($htmlContent, '\sstyle\s*=\s*"[^"]*"', '')

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
