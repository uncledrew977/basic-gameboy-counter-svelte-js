$inputDir = "C:\Path\To\Your\HTML\Files"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    $htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Remove all inline style attributes
    $htmlContent = $htmlContent -replace '\sstyle\s*=\s*"(.*?)"', ''

    # Remove all class attributes
    $htmlContent = $htmlContent -replace '\sclass\s*=\s*"(.*?)"', ''

    # Remove empty inline tags like <span></span>, <div></div>, <font></font>
    $htmlContent = $htmlContent -replace '<(span|div|font)[^>]*>\s*</\1>', ''

    # Remove all HTML comments
    $htmlContent = $htmlContent -replace '<!--.*?-->', ''

    # Remove empty or whitespace-only lines
    $htmlContent = $htmlContent -split "`n" | Where-Object { $_.Trim() -ne "" } | Out-String

    # Save cleaned content
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
