# Load the HTML file
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Raw -Encoding UTF8

# Normalize non-breaking spaces and Word formatting
$html = $html -replace '\u00A0', ' '
$html = $html -replace '&nbsp;', ' '
$html = $html -replace '>\s+<', '><'  # Collapse tag whitespace

# Step 1: Remove all attributes from all HTML tags
$html = [regex]::Replace($html, '<(\w+)(\s[^>]*)?>', '<$1>')

# Step 2: Replace all inner text (between tags) with the letter A
$html = [regex]::Replace($html, '(?<=>)([^<]+)(?=<)', {
    param($m)
    return 'A'
})

# Step 3: Save cleaned HTML
$outputPath = "$env:TEMP\replaced_output.html"
$html | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Cleaned + replaced HTML saved to: $outputPath"
