$inputDir = "C:\Test\Output"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    $htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Remove class attributes (both quote types)
    $htmlContent = [regex]::Replace($htmlContent, '(?i)\sclass\s*=\s*(".*?"|\'.*?\')', '')

    # Remove style attributes that do NOT contain background:silver
    $htmlContent = [regex]::Replace($htmlContent, '(?i)\sstyle\s*=\s*("(?:(?!background\s*:\s*silver)[^"])*?"|\'(?:(?!background\s*:\s*silver)[^\'])*?\')', '')

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
