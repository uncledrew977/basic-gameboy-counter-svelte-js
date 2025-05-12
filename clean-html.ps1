Add-Type -AssemblyName 'System.Web'

$inputDir = "C:\Path\To\Your\HTML\Files"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    [string]$htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Normalize all tag and attribute case (lowercase)
    $htmlContent = $htmlContent -replace '(?i)(<[^>]+>)', { $args[0].ToLower() }

    # Convert <strong>/<em> to <b>/<i>
    $htmlContent = $htmlContent -replace "<strong>", "<b>"
    $htmlContent = $htmlContent -replace "</strong>", "</b>"
    $htmlContent = $htmlContent -replace "<em>", "<i>"
    $htmlContent = $htmlContent -replace "</em>", "</i>"

    # Keep only style="background:silver" and delete all other styles and class/lang
    $htmlContent = $htmlContent -replace 'style="((?:(?!background\s*:\s*silver)[^""])+)"', ''
    $htmlContent = $htmlContent -replace 'style="[^"]*(background\s*:\s*silver)[^"]*"', 'style="$1"'
    $htmlContent = $htmlContent -replace '\sclass="[^"]*"', ''
    $htmlContent = $htmlContent -replace '\slang="[^"]*"', ''

    # Remove Word-specific comments and styles
    $htmlContent = $htmlContent -replace "<!--\[if.*?\]-->", ""
    $htmlContent = $htmlContent -replace "(?s)<style[^>]*>.*?</style>", ""
    $htmlContent = $htmlContent -replace "<meta[^>]*>", ""
    $htmlContent = $htmlContent -replace "mso-[^:]+:[^;""']+;?", ""

    # Remove empty spans, fonts, divs
    $htmlContent = $htmlContent -replace "<(span|font|div)[^>]*>\s*</\1>", ""

    # Remove any remaining empty inline elements
    $htmlContent = $htmlContent -replace "<(span|font|div)[^>]*></\1>", ""

    # Collapse multiple spaces
    $htmlContent = $htmlContent -replace "\s{2,}", " "

    # Write back cleaned content
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
