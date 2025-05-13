# Input and output paths
$htmlPath = "C:\Path\To\Original.html"
$outputPath = "$env:TEMP\redacted_output.html"

# Load HtmlAgilityPack DLL
$hpackDll = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackDll

# Read HTML and wrap
$html = Get-Content $htmlPath -Raw -Encoding Default
$html = "<html><body>" + $html + "</body></html>"

# Load into HtmlAgilityPack
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

# Define bullet detection pattern
$bulletPattern = '^\s*(\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.\)\:]?)\s*'

# Process all text nodes (except inside script/style)
$nodes = $doc.DocumentNode.SelectNodes("//*[not(self::script or self::style)]/text()")
foreach ($textNode in $nodes) {
    $text = $textNode.InnerText
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        # Try to match leading bullet pattern
        if ($text -match $bulletPattern) {
            $bullet = $matches[1]
            $redacted = $text -replace [regex]::Escape($bullet), ''
            $redacted = $redacted -replace '.', 'X'
            $textNode.InnerHtml = "$bullet$redacted"
        } else {
            $textNode.InnerHtml = $text -replace '.', 'X'
        }
    }
}

# Save output
$doc.DocumentNode.InnerHtml | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Redacted HTML (with bullets preserved) saved to: $outputPath"
