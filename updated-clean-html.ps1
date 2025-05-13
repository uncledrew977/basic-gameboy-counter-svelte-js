# Input and output paths
$htmlPath = "C:\Path\To\Original.html"
$outputPath = "$env:TEMP\redacted_output.html"

# Load HtmlAgilityPack DLL
$hpackDll = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackDll

# Read HTML and wrap if needed
$html = Get-Content $htmlPath -Raw -Encoding Default
$html = "<html><body>" + $html + "</body></html>"

# Load HTML
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

# Improved bullet detection (start only, Roman numerals, numbers, dots, symbols)
$bulletPattern = '^\s*(\(?[ivxlcdmIVXLCDM0-9a-zA-Z]{1,5}[\.\)\:]|[•▪·\.])(\s+|&nbsp;)+'

# Select all visible text nodes
$nodes = $doc.DocumentNode.SelectNodes("//*[not(self::script or self::style)]/text()")
foreach ($textNode in $nodes) {
    $text = $textNode.InnerText
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        if ($text -match $bulletPattern) {
            $bulletOnly = $matches[1]  # Only the bullet portion
            $afterBullet = $text -replace '^\s*(\(?[ivxlcdmIVXLCDM0-9a-zA-Z]{1,5}[\.\)\:]|[•▪·\.])(\s+|&nbsp;)+', ''
            $redactedRest = $afterBullet -replace '.', 'X'
            $textNode.InnerHtml = "$bulletOnly $redactedRest"
        } else {
            $textNode.InnerHtml = $text -replace '.', 'X'
        }
    }
}

# Save output
$doc.DocumentNode.InnerHtml | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Redacted HTML saved to: $outputPath"
