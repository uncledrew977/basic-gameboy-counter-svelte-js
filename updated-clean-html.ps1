# Load HtmlAgilityPack
Add-Type -Path "C:\Path\To\HtmlAgilityPack.dll"

# Load the HTML content with correct encoding
$path = "C:\Path\To\Your\File.html"
$html = Get-Content $path -Encoding Default -Raw

# Load the document
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

# Bullet regex: handles 1., a., i., ii., · etc.
$bulletPattern = '^\s*(?:[0-9]+\.|[a-zA-Z]{1,3}\.|[ivxlcdm]+\.)\s*$|^\s*[·•]\s*$'

# Redact logic
function Redact-Text($text) {
    if ($text -match $bulletPattern) {
        return $text
    }
    return ($text -replace '\S', 'X')
}

# Process all text nodes
foreach ($node in $doc.DocumentNode.SelectNodes("//*[not(self::script or self::style)]/text()")) {
    $trimmed = $node.InnerText.Trim()
    if ($trimmed) {
        $node.InnerHtml = Redact-Text $trimmed
    }
}

# Save output
$outputPath = "C:\Path\To\Redacted-Output.html"
$doc.Save($outputPath)
Write-Host "Redacted file saved to $outputPath"
