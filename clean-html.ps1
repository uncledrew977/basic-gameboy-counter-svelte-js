Add-Type -AssemblyName 'System.Web'

$inputDir = "C:\Path\To\Output\HTML"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    [string]$htmlContent = Get-Content $htmlPath -Raw

    # Remove class attributes
    $htmlContent = $htmlContent -replace '\sclass="[^"]*"', ''

    # Remove lang attributes
    $htmlContent = $htmlContent -replace '\slang="[^"]*"', ''

    # Preserve only style attributes that include background:silver
    $htmlContent = $htmlContent -replace 'style="((?![^"]*background\s*:\s*silver)[^"]*)"', ''

    # Optional: normalize background:silver (remove other stuff in style)
    $htmlContent = $htmlContent -replace 'style="[^"]*(background\s*:\s*silver)[^"]*"', 'style="$1"'

    # Remove empty tags (e.g. <span></span>, <font></font>)
    $htmlContent = $htmlContent -replace '<(span|font)[^>]*>\s*</\1>', ''

    # Replace <strong>/<em> with <b>/<i>
    $htmlContent = $htmlContent -replace '<strong>', '<b>' -replace '</strong>', '</b>'
    $htmlContent = $htmlContent -replace '<em>', '<i>' -replace '</em>', '</i>'

    # Remove conditional Word comments and Word styles
    $htmlContent = $htmlContent -replace '<!--\[if.*?endif\]-->', ''
    $htmlContent = $htmlContent -replace '<style[^>]*>.*?</style>', ''
    $htmlContent = $htmlContent -replace '<meta[^>]*>', ''
    $htmlContent = $htmlContent -replace 'mso-[^:]+:[^;"]+;?', ''

    # Trim excess whitespace
    $htmlContent = $htmlContent -replace '\s{2,}', ' '

    # Save cleaned content
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8

    Write-Host "Cleaned: $htmlPath"
}
