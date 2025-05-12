# Load raw HTML and normalize spacing
$inputPath = "C:\Path\To\Test.html"
$html = Get-Content $inputPath -Raw -Encoding UTF8

# Normalize line breaks and spaces (Word spreads tags across lines)
$html = $html -replace '\r?\n', ' '
$html = $html -replace '\u00A0', ' '
$html = $html -replace '&nbsp;', ' '
$html = $html -replace '>\s+<', '><'

# Match all <p> elements
$paragraphs = [regex]::Matches($html, '<p[^>]*?>.*?</p>', 'IgnoreCase')

# Prepare output
$output = ""
$listStack = @()
$indentUnit = 36  # pt per indent level

function Close-Lists($level) {
    while ($listStack.Count -gt $level) {
        $output += "</$($listStack[-1])>`n"
        $listStack = $listStack[0..($listStack.Count - 2)]
    }
}

function Open-List($tag) {
    $output += "<$tag>`n"
    $listStack += $tag
}

foreach ($match in $paragraphs) {
    $p = $match.Value
    $style = if ($p -match 'style="([^"]+)"') { $matches[1] } else { "" }

    # Estimate nesting level from margin-left
    $level = 0
    if ($style -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
        $margin = [double]$matches[1]
        $level = [math]::Round($margin / $indentUnit)
    }

    $isList = $p -match 'MsoListParagraph'

    # Extract raw inner HTML
    $innerHtml = [regex]::Match($p, '>(.*?)</p>', 'Singleline').Groups[1].Value

    if ($isList) {
        # Guess list type from bullet span
        $listType = if ($innerHtml -match '<span[^>]*>[a-zA-Z0-9]+[\.\)\:]') { "ol" } else { "ul" }

        # Strip bullet span (e.g., <span>1.<span>...</span></span>)
        $cleaned = $innerHtml -replace '<span[^>]*>[a-zA-Z0-9]+[\.\)\:]\s*<span[^>]*>.*?</span></span>', ''
        $cleaned = $cleaned -replace '<[^>]+>', ''
        $text = $cleaned.Trim()

        # Manage nesting
        Close-Lists $level
        if ($listStack.Count -lt ($level + 1)) {
            Open-List $listType
        }

        $output += "<li>$text</li>`n"
    }
    else {
        # Not a list, just output as paragraph
        Close-Lists 0
        $text = $innerHtml -replace '<[^>]+>', ''
        $output += "<p>$($text.Trim())</p>`n"
    }
}

# Close any remaining open lists
Close-Lists 0

# Save the result
$outputPath = "$env:TEMP\cleaned_final_output.html"
$output | Set-Content -Encoding UTF8 -Path $outputPath
Write-Output "Final HTML with lists saved to: $outputPath"
