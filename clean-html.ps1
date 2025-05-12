$inputDir = "C:\Test\Output"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    $htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Remove class attributes (quoted or unquoted)
    $htmlContent = [regex]::Replace($htmlContent, '(?i)\sclass\s*=\s*["\']?([^\s>]+)["\']?', '')

    # Remove style attributes (quoted or unquoted)
    $htmlContent = [regex]::Replace($htmlContent, '(?i)\sstyle\s*=\s*["\']?([^\s>]+)["\']?', '')

    # Remove id attributes (quoted or unquoted)
    $htmlContent = [regex]::Replace($htmlContent, '(?i)\sid\s*=\s*["\']?([^\s>]+)["\']?', '')

    # Remove lang attributes (quoted or unquoted)
    $htmlContent = [regex]::Replace($htmlContent, '(?i)\slang\s*=\s*["\']?([^\s>]+)["\']?', '')

    # Remove HTML comments
    $htmlContent = [regex]::Replace($htmlContent, '(?is)<!--.*?-->', '')

    # Remove empty span/div/font tags
    $htmlContent = [regex]::Replace($htmlContent, '(?i)<(span|div|font)[^>]*>\s*</\1>', '')

    # Remove empty lines
    $htmlLines = $htmlContent -split "`n" | Where-Object { $_.Trim() -ne "" }
    $htmlContent = $htmlLines -join "`r`n"

    # Save cleaned HTML
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
