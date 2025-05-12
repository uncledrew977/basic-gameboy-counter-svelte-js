# Load the HTML file
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Raw

# Step 1: Convert Word-generated lists into real HTML lists
function Convert-WordLists {
    param($htmlContent)

    # Extract <p> tags that look like list items (with class or margin-left style)
    $pattern = '<p[^>]*(MsoListParagraph|margin-left:[^">]+)[^>]*>(.*?)</p>'
    $matches = [regex]::Matches($htmlContent, $pattern, 'IgnoreCase')

    $output = ''
    $listStack = @()
    $prevLevel = 0

    foreach ($match in $matches) {
        $fullTag = $match.Value
        $innerHtml = $match.Groups[2].Value.Trim()

        if ($fullTag -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
            $margin = [double]$matches[1].Groups[1].Value
            $level = [math]::Round($margin / 36)
        } else {
            $level = 0
        }

        if ($innerHtml -match '^\s*<span[^>]*>[•·▪◦]') {
            $listType = 'ul'
        } elseif ($innerHtml -match '^\s*<span[^>]*>\s*[\diIvVaA]+\.' ) {
            $listType = 'ol'
        } else {
            $listType = 'ul'
        }

        $itemText = $innerHtml -replace '^<span[^>]*>.*?</span>', ''
        $itemText = $itemText.Trim()

        while ($prevLevel < $level) {
            $output += "<$listType>"
            $listStack += $listType
            $prevLevel++
        }

        while ($prevLevel > $level) {
            $lastList = $listStack[-1]
            $output += "</$lastList>"
            $listStack = $listStack[0..($listStack.Count - 2)]
            $prevLevel--
        }

        $output += "<li>$itemText</li>"
    }

    while ($listStack.Count -gt 0) {
        $lastList = $listStack[-1]
        $output += "</$lastList>"
        $listStack = $listStack[0..($listStack.Count - 2)]
    }

    $htmlContent = $htmlContent -replace $pattern, ''
    return $htmlContent + $output
}

# Run list converter
$html = Convert-WordLists $html

# Remove class=""
$html = $html -replace '\sclass=""', ''

# Remove style unless it contains background:silver
$html = [regex]::Replace($html, '\sstyle="([^"]*)"', {
    param($match)
    $style = $match.Groups[1].Value
    if ($style -match 'background\s*:\s*silver') {
        return $match.Value
    } else {
        return ''
    }
})

# Strip attributes inside <body> but preserve background:silver style
function Strip-Attributes-InBody {
    param($htmlContent)

    if ($htmlContent -match '(?s)(<body[^>]*>)(.*?)(</body>)') {
        $bodyOpen = $matches[1]
        $bodyContent = $matches[2]
        $bodyClose = $matches[3]

        $cleanBodyContent = [regex]::Replace($bodyContent, '<(\w+)([^>]*)>', {
            param($tagMatch)
            $tag = $tagMatch.Groups[1].Value
            $attrs = $tagMatch.Groups[2].Value

            if ($attrs -match 'style\s*=\s*"([^"]*background\s*:\s*silver[^"]*)"\s*') {
                $style = $matches[1]
                return "<$tag style=`"$style`">"
            } else {
                return "<$tag>"
            }
        })

        return $htmlContent -replace [regex]::Escape($matches[0]), "$bodyOpen$cleanBodyContent$bodyClose"
    } else {
        return $htmlContent
    }
}

# Clean body
$html = Strip-Attributes-InBody $html

# Save with Set-Content
$outputPath = "$env:TEMP\cleaned_output.html"
$html | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Cleaned HTML saved to: $outputPath"
