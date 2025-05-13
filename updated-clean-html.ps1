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

# Bullet detection pattern (includes Roman numerals, dots, bullets)
$bulletPattern = '^\s*((\(?[ivxlcdmIVXLCDM0-9a-zA-Z]{1,5}[\.\)\:]|[•▪·\.]))(\s+|(&nbsp;)+)'

# Process all non-script/style text nodes
$nodes = $doc.DocumentNode.SelectNodes("//*[not(self::script or self::style)]/text()")
foreach ($textNode in $nodes) {
    $text = $textNode.InnerText
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        if ($text -match $bulletPattern) {
            $bullet = $matches[1]
            $redactedRest = $text.Substring($matches[0].Length) -replace '.', 'X'
            $textNode.InnerHtml = "$bullet$redactedRest"
        } else {
            $textNode.InnerHtml = $text -replace '.', 'X'
        }
    }
}

# Save redacted output
$doc.DocumentNode.InnerHtml | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Redacted HTML (with Roman numerals and bullets preserved) saved to: $outputPath"
