# Load the HTML file
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Raw -Encoding UTF8

# Normalize Word artifacts
$html = $html -replace '\u00A0', ' '
$html = $html -replace '&nbsp;', ' '
$html = $html -replace '>\s+<', '><'  # Collapse whitespace between tags

# Step 1: DO NOT strip attributes (leave tags untouched)

# Step 2: Replace all inner text with "X", except bullet-like markers
$bulletPattern = '^\s*(\(?[a-z]{1,5}[\.\)\:]|\(?[ivxlcdm]{1,5}[\.\)\:]|\(?\d{1,4}[\.\)\:])\s*$'

$html = [regex]::Replace($html, '(?<=>)([^<]+)(?=<)', {
    param($m)
    $text = $m.Groups[1].Value.Trim()
    if ($text -match $bulletPattern) {
        return $text  # Preserve bullet markers
    } else {
        return 'X'
    }
})

# Save the output for inspection
$outputPath = "$env:TEMP\test_output_with_attributes.html"
$html | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Output with original attributes saved to: $outputPath"
