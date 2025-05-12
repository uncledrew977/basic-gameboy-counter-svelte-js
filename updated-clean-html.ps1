# Load the HTML file with UTF-8 encoding (handles BOM)
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Encoding UTF8 -Raw

# Normalize non-breaking spaces
$html = $html -replace '\u00A0', ' '
$html = $html -replace '&nbsp;', ' '

# Step 1: Convert Word-style <p> list blocks into real HTML lists
function Convert-WordLists {
    param($htmlContent)

    $pattern = '(?s)<p[^>]*?(MsoListParagraph|text-indent:-\.25in|margin-left:\s*\d+pt)[^>]*?>(.*?)</p>'
    $matches = [regex]::Matches($htmlContent, $pattern)

    $output = ''
    $listStack = @()
    $prevLevel = 0

    foreach ($match in $matches) {
        $fullTag = $match.Value
        $innerHtml = $match.Groups[2].Value

        # Determine nesting level from margin-left
        if ($fullTag -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
            $marginMatch = [regex]::Match($fullTag, 'margin-left:\s*(\d+(\.\d+)?)pt')
            $margin = [double]$marginMatch.Groups[1].Value
            $level = [math]::Round($margin / 18)
        } else {
            $level = 0
        }

        # Determine list type (ordered vs unordered)
        $bulletMatch = [regex]::Match($innerHtml, '<span[^>]*>([a-zA-Z0-9]+)[\.\)]\s*<span[^>]*>.*?</span></span>')
        $bullet = $bulletMatch.Groups[1].Value

        if ($bullet -match '^\d+$' -or $bullet -match '^[a-zA-Z]+$') {
            $listType = 'ol'
        } else {
            $listType = 'ul'
        }

        # Extract the actual list content (remove all tags except the text after the bullet)
        $text = $innerHtml -replace '<span[^>]*>[a-zA-Z0-9]+\.\s*<span[^>]*>.*?</span></span>', ''
        $text = $text -replace '<[^>]+>', ''
        $itemText = $text.Trim()

        # Handle nesting
        while ($prevLevel -lt $level) {
            $output += ("<" + $listType + ">")
            $listStack += $listType
            $prevLevel++
        }

        while ($prevLevel -gt $level -and $listStack.Count -gt 0) {
            $lastList = $listStack[-1]
            $output += ("</" + $lastList + ">")
            if ($listStack.Count -gt 1) {
                $listStack = $listStack[0..($listStack.Count - 2)]
            } else {
                $listStack = @()
            }
            $prevLevel--
        }

        $output += ("<li>" + $itemText + "</li>")
    }

    while ($listStack.Count -gt 0) {
        $lastList = $listStack[-1]
        $output += ("</" + $lastList + ">")
        if ($listStack.Count -gt 1) {
            $listStack = $listStack[0..($listStack.Count - 2)]
        } else {
            $listStack = @()
        }
    }

    # Remove original <p> blocks
    $htmlContent = $htmlContent -replace $pattern, ''
    return $htmlContent + $output
}

# Step 2: Convert Word-style lists to real HTML lists
$html = Convert-WordLists $html

# Step 3: Remove all attributes inside <body> (completely clean tags)
function Strip-AllAttributes-InBody {
    param($htmlContent)

    if ($htmlContent -match '(?s)(<body[^>]*>)(.*?)(</body>)') {
        $bodyOpen = $matches[1]
        $bodyContent = $matches[2]
        $bodyClose = $matches[3]

        # Remove all attributes from all tags inside <body>
        $cleanBodyContent = [regex]::Replace($bodyContent, '<(\w+)(\s[^>]*)?>', '<$1>')

        return $htmlContent -replace [regex]::Escape($matches[0]), ($bodyOpen + $cleanBodyContent + $bodyClose)
    } else {
        return $htmlContent
    }
}

# Step 4: Clean <body> attributes
$html = Strip-AllAttributes-InBody $html

# Step 5: Save final cleaned HTML
$outputPath = "$env:TEMP\cleaned_output.html"
$html | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Cleaned HTML saved to: $outputPath"
