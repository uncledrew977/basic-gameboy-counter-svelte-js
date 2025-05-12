# Load HTML file
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Raw -Encoding UTF8

# Normalize non-breaking spaces
$html = $html -replace '\u00A0', ' '
$html = $html -replace '&nbsp;', ' '

# Step 0: Normalize line breaks inside tags (collapse Office formatting)
$html = $html -replace '>\s+<', '><'
$html = $html -replace '\r?\n', ' '

# Step 1: Convert Word-style list <p> blocks to <ol>/<ul>/<li>
function Convert-WordLists {
    param($htmlContent)

    # Match <p> blocks that look like Word lists (class or indent-based)
    $pattern = '(?s)<p[^>]*?(MsoListParagraph|text-indent:-?\.?\d+in|margin-left:\s*\d+pt)[^>]*?>(.*?)</p>'
    $matches = [regex]::Matches($htmlContent, $pattern)

    $output = ''
    $listStack = @()
    $prevLevel = 0

    foreach ($match in $matches) {
        $fullTag = $match.Value
        $innerHtml = $match.Groups[2].Value

        # Determine indentation level
        if ($fullTag -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
            $margin = [double]([regex]::Match($fullTag, 'margin-left:\s*(\d+(\.\d+)?)pt').Groups[1].Value)
            $level = [math]::Round($margin / 18)
        } else {
            $level = 0
        }

        # Determine list type (ordered if bullet-like prefix exists)
        if ($innerHtml -match '<span[^>]*>[a-zA-Z0-9]+[\.\)]') {
            $listType = 'ol'
        } else {
            $listType = 'ul'
        }

        # Remove bullet span
        $content = $innerHtml -replace '<span[^>]*>[a-zA-Z0-9]+[\.\)]\s*(<span[^>]*>.*?</span>)?</span>', ''
        # Remove all remaining tags to extract clean text
        $cleanText = $content -replace '<[^>]+>', ''
        $itemText = $cleanText.Trim()

        # Nesting logic
        while ($prevLevel -lt $level) {
            $output += "<$listType>"
            $listStack += $listType
            $prevLevel++
        }

        while ($prevLevel -gt $level -and $listStack.Count -gt 0) {
            $lastList = $listStack[-1]
            $output += "</$lastList>"
            $listStack = if ($listStack.Count -gt 1) { $listStack[0..($listStack.Count - 2)] } else { @() }
            $prevLevel--
        }

        $output += "<li>$itemText</li>"
    }

    # Close remaining lists
    while ($listStack.Count -gt 0) {
        $lastList = $listStack[-1]
        $output += "</$lastList>"
        $listStack = if ($listStack.Count -gt 1) { $listStack[0..($listStack.Count - 2)] } else { @() }
    }

    # Remove original list <p> tags
    foreach ($match in $matches) {
        $htmlContent = $htmlContent -replace [regex]::Escape($match.Value), ''
    }

    return $htmlContent + $output
}

# Step 2: Run list converter
$html = Convert-WordLists $html

# Step 3: Remove ALL attributes from ALL tags inside <body>
function Strip-AllAttributes-InBody {
    param($htmlContent)

    if ($htmlContent -match '(?s)(<body[^>]*>)(.*?)(</body>)') {
        $bodyOpen = $matches[1]
        $bodyContent = $matches[2]
        $bodyClose = $matches[3]

        # Strip all attributes (e.g., <span style="..."> => <span>)
        $cleanBodyContent = [regex]::Replace($bodyContent, '<(\w+)(\s[^>]*)?>', '<$1>')

        return $htmlContent -replace [regex]::Escape($matches[0]), ($bodyOpen + $cleanBodyContent + $bodyClose)
    } else {
        return $htmlContent
    }
}

# Step 4: Strip all attributes
$html = Strip-AllAttributes-InBody $html

# Step 5: Save cleaned output
$outputPath = "$env:TEMP\cleaned_output.html"
$html | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Cleaned HTML saved to: $outputPath"
