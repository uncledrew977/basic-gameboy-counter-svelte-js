Add-Type -AssemblyName 'System.Web'

$inputDir = "C:\Path\To\Your\HTML\Files"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    [string]$htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Convert <strong>/<em> to <b>/<i>
    $htmlContent = $htmlContent -replace "(?i)<\s*strong\s*>", "<b>"
    $htmlContent = $htmlContent -replace "(?i)<\s*/\s*strong\s*>", "</b>"
    $htmlContent = $htmlContent -replace "(?i)<\s*em\s*>", "<i>"
    $htmlContent = $htmlContent -replace "(?i)<\s*/\s*em\s*>", "</i>"

    # Remove all class and lang attributes (case-insensitive)
    $htmlContent = $htmlContent -replace "(?i)\sclass\s*=\s*(""[^""]*""|'[^']*')", ""
    $htmlContent = $htmlContent -replace "(?i)\slang\s*=\s*(""[^""]*""|'[^']*')", ""

    # Remove style unless it contains background:silver
    $htmlContent = $htmlContent -replace '(?i)style\s*=\s*"(?:(?!background\s*:\s*silver)[^""])*"', ""
    $htmlContent = $htmlContent -replace '(?i)style="[^"]*(background\s*:\s*silver)[^"]*"', 'style="$1"'

    # Remove Office-specific styles/comments
    $htmlContent = $htmlContent -replace "(?i)<!--\[if.*?endif\]-->", ""
    $htmlContent = $htmlContent -replace "(?is)<style[^>]*>.*?</style>", ""
    $htmlContent = $htmlContent -replace "(?i)<meta[^>]*>", ""
    $htmlContent = $htmlContent -replace "(?i)mso-[^:]+:[^;""']+;?", ""

    # Remove empty spans, fonts, and divs
    $htmlContent = $htmlContent -replace "(?i)<(span|font|div)[^>]*>\s*</\1>", ""

    # Preserve <ol>, <ul>, <li> (with type attributes for roman/lettered/numbered)
    # Do NOT touch any <ol.*?>, <ul.*?>, or <li> elements

    # Collapse excessive whitespace
    $htmlContent = $htmlContent -replace "\s{2,}", " "

    # Save cleaned HTML
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
