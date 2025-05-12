# Input and output paths
$inputPath = "C:\Path\To\Test.html"
$outputPath = "$env:TEMP\cleaned_lists.html"

# Read HTML
$html = Get-Content $inputPath -Raw -Encoding UTF8

# Normalize line breaks and spaces
$html = $html -replace '\u00A0', ' '
$html = $html -replace '&nbsp;', ' '
$html = $html -replace '>\s+<', '><'

# Regex to find list-like <p> blocks
$pattern = '<p[^>]*?((MsoListParagraph)|margin-left:\s*\d+pt|text-indent)[^>]*?>(.*?)</p>'

# Match all <p> blocks (including non-lists to preserve them)
$allParagraphs = [regex]::Matches($html, '<p[^>]*?>.*?</p>', 'IgnoreCase')

# Indentation unit to estimate nesting (Word uses ~36pt per level)
$indentUnit = 36
$listStack = @()
$output = ""

function Close-Lists {
    param([int]$targetLevel)
    while ($listStack.Count -gt $targetLevel) {
        $tag = $listStack[-1]
        $script:output += "</$tag>`n"
        $listStack = $listStack[0..($listStack.Count - 2)]
    }
}

function Open-List {
    param($type)
    $script:output += "<$type>`n"
    $listStack += $type
}

foreach ($paraMatch in $allParagraphs) {
    $p = $paraMatch.Value
    $style = if ($p -match 'style="([^"]+)"') { $matches[1] } else { "" }

    $level = 0
    if ($style -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
        $margin = [double]$matches[1]
        $level = [math]::Round($margin / $indentUnit)
    }

    $isList = $p -match 'MsoListParagraph|margin-left:\s*\d+pt|text-indent'
    $innerHtml = [regex]::Match($p, '>(.*?)</p>', 'Singleline').Groups[1].Value

    if ($isList) {
        # Detect ordered vs unordered by bullet shape
        if ($innerHtml -match '<span[^>]*>[a-zA-Z0-9]+[\.\)\:]') {
            $listType = "ol"
        } else {
            $listType = "ul"
        }

        # Strip bullet prefix span
        $cleaned = [regex]::Replace($innerHtml, '<span[^>]*>[a-zA-Z0-9]+[\.\)\:](.*?)</span>', '', 'IgnoreCase')
        $textOnly = [regex]::Replace($cleaned, '<[^>]+>', '').Trim()

        # Open/close lists based on nesting level
        Close-Lists -targetLevel:$level
        if ($listStack.Count -lt ($level + 1)) {
            Open-List -type:$listType
        }

        $output += "<li>$textOnly</li>`n"
    }
    else {
        # Non-list <p> element: close all lists first
        Close-Lists -targetLevel:0
        $text = [regex]::Replace($innerHtml, '<[^>]+>', '').Trim()
        $output += "<p>$text</p>`n"
    }
}

# Close any remaining open lists
Close-Lists -targetLevel:0

# Write to file
Set-Content -Path $outputPath -Value $output -Encoding UTF8
Write-Output "Cleaned HTML with lists saved to: $outputPath"
