Add-Type -AssemblyName 'System.Web'

$inputDir = "C:\Path\To\Your\HTML\Files"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    [string]$htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Normalize tag case by using a regex MatchEvaluator
    $htmlContent = [System.Text.RegularExpressions.Regex]::Replace(
        $htmlContent,
        '<[^>]+>',
        { param($match) $match.Value.ToLower() }
    )

    # Convert <strong>/<em> to <b>/<i>
    $htmlContent = $htmlContent -replace "<strong>", "<b>"
    $htmlContent = $htmlContent -replace "</strong>", "</b>"
    $htmlContent = $htmlContent -replace "<em>", "<i>"
    $htmlContent = $htmlContent -replace "</em>", "</i>"

    # Remove all style attributes except background:silver
    $htmlContent = $htmlContent -replace 'style="((?:(?!background\s*:\s*silver)[^""])+)"', ''
    $htmlContent = $htmlContent -replace 'style="[^"]*(background\s*:\s*silver)[^"]*"', 'style="$1"'

    # Remove class and lang attributes (case-normalized already)
    $htmlContent = $htmlContent -replace '\sclass="[^"]*"', ''
    $htmlContent = $htmlContent -replace '\slang="[^"]*"', ''

    # Remove Word-specific junk
    $htmlContent = $htmlContent -replace "<!--\[if.*?\]-->", ""
    $htmlContent = $htmlContent -replace "(?s)<style[^>]*>.*?</style>", ""
    $htmlContent = $htmlContent -replace "<meta[^>]*>", ""
    $htmlContent = $htmlContent -replace "mso-[^:]+:[^;""']+;?", ""

    # Remove empty inline elements
    $htmlContent = $htmlContent -replace "<(span|font|div)[^>]*>\s*</\1>", ""
    $htmlContent = $htmlContent -replace "<(span|font|div)[^>]*></\1>", ""

    # Collapse extra spaces
    $htmlContent = $htmlContent -replace "\s{2,}", " "

    # Save cleaned HTML
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
