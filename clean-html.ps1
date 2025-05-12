$inputDir = "C:\Path\To\Your\HTML\Files"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    $htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Remove all inline style attributes (double or single quoted, any spacing)
    $htmlContent = $htmlContent -replace '(?i)\s*style\s*=\s*(".*?"|\'.*?\')', ''

    # Remove all class attributes (double or single quoted, any spacing)
    $htmlContent = $htmlContent -replace '(?i)\s*class\s*=\s*(".*?"|\'.*?\')', ''

    # Remove empty tags like <span></span>, <div></div>, <font></font>
    $htmlContent = $htmlContent -replace '(?i)<(span|div|font)[^>]*>\s*</\1>', ''

    # Remove HTML comments
    $htmlContent = $htmlContent -replace '(?s)<!--.*?-->', ''

    # Remove blank lines
    $htmlContent = $htmlContent -split "`n" | Where-Object { $_.Trim() -ne "" } | Out-String

    # Save cleaned HTML
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
