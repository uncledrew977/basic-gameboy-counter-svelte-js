# Path to the HTML file
$htmlPath = "C:\Path\To\Test.html"
$outputPath = "$env:TEMP\cleaned_with_lists.html"

# Load HtmlAgilityPack
$hpackDll = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\netstandard2.0\HtmlAgilityPack.dll"
Add-Type -Path $hpackDll

# Load HTML content
$html = Get-Content $htmlPath -Raw -Encoding UTF8
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

# Constants
$indentUnit = 36
$listStack = @()

function Close-Lists($level, [ref]$outputBuilder) {
    while ($listStack.Count -gt $level) {
        $closing = "</$($listStack[-1])>`n"
        $outputBuilder.Value += $closing
        $listStack = $listStack[0..($listStack.Count - 2)]
    }
}

function Open-List($type, [ref]$outputBuilder) {
    $outputBuilder.Value += "<$type>`n"
    $listStack += $type
}

# Builder for output HTML
$outputBuilder = [ref] ''

# Go through all top-level nodes
foreach ($node in $doc.DocumentNode.SelectNodes("//*")) {
    if ($node.Name -eq "p" -and $node.GetAttributeValue("class", "") -like "*MsoListParagraph*") {
        $style = $node.GetAttributeValue("style", "")
        $level = 0
        if ($style -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
            $margin = [double]$matches[1]
            $level = [math]::Round($margin / $indentUnit)
        }

        # Extract text, remove bullets
        $text = $node.InnerText.Trim()
        if ($text -match '^[a-zA-Z0-9]{1,5}[\.\)\:]\s*') {
            $text = $text -replace '^[a-zA-Z0-9]{1,5}[\.\)\:]\s*', ''
        }

        # Determine list type
        $listType = if ($node.InnerHtml -match '[a-zA-Z0-9]{1,5}[\.\)\:]') { "ol" } else { "ul" }

        Close-Lists $level $outputBuilder
        if ($listStack.Count -lt ($level + 1)) {
            Open-List $listType $outputBuilder
        }

        $outputBuilder.Value += "<li>$text</li>`n"
        $node.Remove()  # Remove from final output
    }
}

# Close remaining open lists
Close-Lists 0 $outputBuilder

# Append remaining HTML after cleaned lists
$remaining = $doc.DocumentNode.InnerHtml
$outputBuilder.Value += "`n$remaining"

# Save output
Set-Content -Path $outputPath -Value $outputBuilder.Value -Encoding UTF8
Write-Output "Cleaned HTML with lists saved to: $outputPath"
