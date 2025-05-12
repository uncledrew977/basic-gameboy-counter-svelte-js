Add-Type -AssemblyName 'System.Web'

$inputDir = "C:\Path\To\Output\HTML"

Get-ChildItem -Path $inputDir -Filter *.html | ForEach-Object {
    $htmlPath = $_.FullName
    [string]$htmlContent = Get-Content $htmlPath -Raw

    # Use regex to clean Word garbage

    # 1. Strip all class attributes
    $htmlContent = $htmlContent -replace '\sclass="[^"]*"', ''

    # 2. Strip all lang attributes
    $htmlContent = $htmlContent -replace '\slang="[^"]*"', ''

    # 3. Strip all style attributes except those with background:silver
    $htmlContent = $htmlContent -replace 'style="(?![^"]*background\s*:\s*silver)[^"]*"', ''

    # 4. Remove empty spans, fonts, etc.
    $htmlContent = $htmlContent -replace '<(span|font)[^>]*>\s*</\1>', ''

    # 5. Ensure bold and italic tags are preserved (replace <strong> with <b>, <em> with <i>)
    $htmlContent = $htmlContent -replace '<strong>', '<b>' -replace '</strong>', '</b>'
    $htmlContent = $htmlContent -replace '<em>', '<i>' -replace '</em>', '</i>'

    # 6. Optional: remove meta and mso styles (from Word)
    $htmlContent = $htmlContent -replace '<!--\[if.*?endif\]-->', '' -replace '<meta[^>]*>', ''
    $htmlContent = $htmlContent -replace '<style[^>]*>.*?</style>', '' -replace 'mso-[^:]+:[^;"]+;?', ''

    # 7. Optional: cleanup leftover spaces/tags
    $htmlContent = $htmlContent -replace '\s{2,}', ' '

    # Save the cleaned HTML back to file
    Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8

    Write-Host "Cleaned: $htmlPath"
}
