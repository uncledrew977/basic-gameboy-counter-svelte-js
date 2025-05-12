Add-Type -AssemblyName 'System.Web'

$inputDir = "C:\Path\To\Your\HTML\Files"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    [string]$htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Convert <strong>/<em> to <b>/<i>
    $htmlContent = $htmlContent -replace "<\s*strong\s*>", "<b>"
    $htmlContent = $htmlContent -replace "<\s*/\s*strong\s*>", "</b>"
    $htmlContent = $htmlContent -replace "<\s*em\s*>", "<i>"
    $htmlContent = $htmlContent -replace "<\s*/\s*em\s*>", "</i>"

    # Remove class and lang attributes
    $htmlContent = $htmlContent -replace "\sclass\s*=\s*(""[^""]*""|'[^']*')", ""
    $htmlContent = $htmlContent -replace "\slang\s*=\s*(""[^""]*""|'[^']*')", ""

    # Keep only style="background:silver", remove all other styles
    $htmlContent = $htmlContent -replace "style\s*=\s*""(?![^""]*background\s*:\s*silver)[^""]*""", ""
    $htmlContent = $htmlContent -replace "style=""[^""]*(background\s*:\s*silver)[^""]*""", 'style="$1"'

    # Remove Office-specific formatting
    $htmlContent = $htmlContent -replace "<!--\[if.*?endif\]-->", ""
    $htmlContent = $htmlContent -replace "<style[^>]*>.*?</style>", ""
    $htmlContent = $htmlContent -replace "<meta[^>]*>", ""
    $htmlContent = $htmlContent -replace "mso-[^:]+:[^;""']+;?", ""

    # Remove empty spans, fonts, divs
    $htmlContent = $htmlContent -replace "<(span|font|div)[^>]*>\s*</\1>", ""

    # Preserve nested and lettered lists — do NOT remove <ul>, <ol>, <li>, etc.

    # Collapse excessive whitespace
    $htmlContent = $htmlContent -replace "\s{2,}", " "

    # Save cleaned HTML
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
