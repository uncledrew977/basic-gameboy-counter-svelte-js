Add-Type -AssemblyName 'System.Web'

$inputDir = "C:\Path\To\Your\HTML\Files"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    [string]$htmlContent = Get-Content $htmlPath -Raw

    Write-Host "Cleaning: $htmlPath"

    # Remove all class attributes (single or double quotes)
    $htmlContent = $htmlContent -replace '\sclass\s*=\s*(".*?"|\'.*?\')', ''

    # Remove all lang attributes
    $htmlContent = $htmlContent -replace '\slang\s*=\s*(".*?"|\'.*?\')', ''

    # Remove style attributes unless they contain background:silver
    $htmlContent = $htmlContent -replace 'style\s*=\s*"(?![^"]*background\s*:\s*silver)[^"]*"', ''

    # Reduce style="... background:silver ..." to only background:silver
    $htmlContent = $htmlContent -replace 'style="[^"]*(background\s*:\s*silver)[^"]*"', 'style="$1"'

    # Replace <strong>/<em> with <b>/<i>
    $htmlContent = $htmlContent -replace '<\s*strong\s*>', '<b>'
    $htmlContent = $htmlContent -replace '<\s*/\s*strong\s*>', '</b>'
    $htmlContent = $htmlContent -replace '<\s*em\s*>', '<i>'
    $htmlContent = $htmlContent -replace '<\s*/\s*em\s*>', '</i>'

    # Remove empty spans, fonts, divs
    $htmlContent = $htmlContent -replace '<(span|font|div)[^>]*>\s*</\1>', ''

    # Remove Word-specific conditional comments
    $htmlContent = $htmlContent -replace '<!--\[if.*?endif\]-->', ''

    # Remove Word <style> blocks
    $htmlContent = $htmlContent -replace '<style[^>]*>.*?</style>', ''

    # Remove meta tags
    $htmlContent = $htmlContent -replace '<meta[^>]*>', ''

    # Remove mso- styles (Microsoft Office)
    $htmlContent = $htmlContent -replace 'mso-[^:]+:[^;"]+;?', ''

    # Clean up whitespace
    $htmlContent = $htmlContent -replace '\s{2,}', ' '

    # Save cleaned HTML
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
}
