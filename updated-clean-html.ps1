# Input and output paths
$htmlPath = "C:\Path\To\Original.html"
$outputPath = "$env:TEMP\redacted_output.html"

# Load HtmlAgilityPack DLL
$hpackDll = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackDll

# Read HTML
$html = Get-Content $htmlPath -Raw -Encoding Default
$html = "<html><body>" + $html + "</body></html>"  # Wrap if needed

# Load into HtmlAgilityPack
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

# Replace all inner text with Xs (skip script/style tags)
$nodes = $doc.DocumentNode.SelectNodes("//*[not(self::script or self::style)]/text()")
foreach ($textNode in $nodes) {
    $original = $textNode.InnerText
    $length = $original.Length
    if ($length -gt 0) {
        $redacted = 'X' * $length
        $textNode.InnerHtml = $redacted
    }
}

# Save output
$doc.DocumentNode.InnerHtml | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Redacted HTML saved to: $outputPath"
