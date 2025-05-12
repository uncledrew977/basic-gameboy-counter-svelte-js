# Load the HTML file
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Raw -Encoding UTF8

# Normalize common spacing artifacts
$html = $html -replace '\u00A0', ' '
$html = $html -replace '&nbsp;', ' '
$html = $html -replace '>\s+<', '><'  # collapse whitespace between tags

# Step 1: Remove all attributes from tags
$html = [regex]::Replace($html, '<(\w+)(\s[^>]*)?>', '<$1>')

# Step 2: Replace all text nodes with "X", except bullet-like text
# Bullet pattern: 1., 2), a., b), i., ii), etc.
$bulletPattern = '^\s*(\(?[a-z]{1,5}[\.\)\:]|\(?[ivxlcdm]{1,5}[\.\)\:]|\(?\d{1,4}[\.\)\:])\s*$'

$html = [regex]::Replace($html, '(?<=>)([^<]+)(?=<)', {
    param($m)
    $text = $m.Groups[1].Value.Trim()
    if ($text -match $bulletPattern) {
        return $text  # leave bullet markers alone
    } else {
        return 'X'
    }
})

# Save the output
$outputPath = "$env:TEMP\replaced_output.html"
$html | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "HTML with replaced text saved to: $outputPath"
