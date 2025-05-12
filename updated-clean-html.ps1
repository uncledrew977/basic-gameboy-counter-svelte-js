# Load the HTML file (UTF-8 handles BOM safely)
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Encoding UTF8 -Raw

# Normalize non-breaking spaces
$html = $html -replace '\u00A0', ' '
$html = $html -replace '&nbsp;', ' '

# Step 1: Convert Word-style list <p> elements to real HTML lists
function Convert-WordLists {
    param($htmlContent)

    # Match <p> blocks with list styling
    $pattern = '(?s)<p[^>]*?(MsoListParagraph|text-indent:-?\.?\d+in|margin-left:\s*\d+pt)[^>]*?>(.*?)</p>'
    $matches = [regex]::Matches($htmlContent, $pattern)

    $output = ''
    $listStack = @()
    $prevLevel = 0

    foreach ($match in $matches) {
        $fullTag = $match.Value
        $innerHtml = $match.Groups[2].Value

        # Determine nesting level based on margin-left
        if ($fullTag -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
            $margin = [double]([regex]::Match($fullTag, 'margin-left:\s*(\d+(\.\d+)?)pt').Groups[1].Value)
            $level = [math]::Round($margin / 18)
        } else {
            $level = 0
        }

        # Determine list type (ol if it starts with a known marker like '1.', 'a.', 'i.')
        if ($innerHtml -match '<span[^>]*>[a-zA-Z0-9]+[\.\)]') {
            $listType = 'ol'
        } else {
            $listType = 'ul'
        }

        # Strip bullet prefix spans and clean text
        $textWithoutBullet = $innerHtml -replace '<span[^>]*>[a-zA-Z0-9]+[\.\)]\s*<span[^>]*>.*?</span></span>', ''
        $textOnly = $textWithoutBullet -replace '<[^>]+>', ''
        $itemText = $textOnly.Trim()

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

    # Close any remaining open lists
    while ($listStack.Count -gt 0) {
        $lastList = $listStack[-1]
        $output += "</$lastList>"
        $listStack = if ($listStack.Count -gt 1) { $listStack[0..($listStack.Count - 2)] } else { @() }
    }

    # Remove only matched <p> list blocks from original HTML
    foreach ($match in $matches) {
        $htmlContent = $htmlContent -replace [regex]::Escape($match.Value), ''
    }

    return $htmlContent + $output
}

# Step 2: Convert lists
$html = Convert-WordLists $html

# Step 3: Remove all attributes from elements inside <body>
function Strip-AllAttributes-InBody {
    param($htmlContent)

    if ($htmlContent -match '(?s)(<body[^>]*>)(.*?)(</body>)') {
        $bodyOpen = $matches[1]
        $bodyContent = $matches[2]
        $bodyClose = $matches[3]

        $cleanBodyContent = [regex]::Replace($bodyContent, '<(\w+)(\s[^>]*)?>', '<$1>')
        return $htmlContent -replace [regex]::Escape($matches[0]), ($bodyOpen + $cleanBodyContent + $bodyClose)
    } else {
        return $htmlContent
    }
}

# Step 4: Clean body tag attributes
$html = Strip-AllAttributes-InBody $html

# Step 5: Save cleaned HTML
$outputPath = "$env:TEMP\cleaned_output.html"
$html | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Cleaned HTML saved to: $outputPath"
