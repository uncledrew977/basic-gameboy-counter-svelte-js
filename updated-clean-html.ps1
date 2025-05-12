# Load the HTML file with UTF-8 BOM support
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Encoding UTF8 -Raw

# Optional: normalize Word's non-breaking spaces
$html = $html -replace '\u00A0', ' '
$html = $html -replace '&nbsp;', ' '

# Step 1: Convert Word-style list <p> blocks into nested <ol>/<ul><li>
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

        # Determine nesting level based on margin-left
        if ($fullTag -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
            $marginMatch = [regex]::Match($fullTag, 'margin-left:\s*(\d+(\.\d+)?)pt')
            $margin = [double]$marginMatch.Groups[1].Value
            $level = [math]::Round($margin / 18)  # 18pt per level
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

        # Extract content from span with background:silver
        $itemMatch = [regex]::Match($innerHtml, '<span[^>]*background:silver[^>]*>(.*?)</span>', 'IgnoreCase')
        $itemText = if ($itemMatch.Success) { $itemMatch.Groups[1].Value.Trim() } else { 'UNKNOWN' }

        # Handle nesting levels
        while ($prevLevel -lt $level) {
            $output += ("<" + $listType + ">")
            $listStack += $listType
            $prevLevel++
        }

        while ($prevLevel -gt $level) {
            $lastList = $listStack[-1]
            $output += ("</" + $lastList + ">")
            $listStack = $listStack[0..($listStack.Count - 2)]
            $prevLevel--
        }

        $output += ("<li>" + $itemText + "</li>")
    }

    while ($listStack.Count -gt 0) {
        $lastList = $listStack[-1]
        $output += ("</" + $lastList + ">")
        $listStack = $listStack[0..($listStack.Count - 2)]
    }

    # Remove the original Word-style <p> blocks
    $htmlContent = $htmlContent -replace $pattern, ''
    return $htmlContent + $output
}

# Step 2: Convert lists before any other cleaning
$html = Convert-WordLists $html

# Step 3: Remove empty class=""
$html = $html -replace '\sclass=""', ''

# Step 4: Remove style unless it contains background:silver
$html = [regex]::Replace($html, '\sstyle="([^"]*)"', {
    param($match)
    $style = $match.Groups[1].Value
    if ($style -match 'background\s*:\s*silver') {
        return $match.Value
    } else {
        return ''
    }
})

# Step 5: Remove all other attributes inside <body>, preserve background:silver style
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
                return ("<" + $tag + " style=`"" + $style + "`">")
            } else {
                return ("<" + $tag + ">")
            }
        })

        return $htmlContent -replace [regex]::Escape($matches[0]), ($bodyOpen + $cleanBodyContent + $bodyClose)
    } else {
        return $htmlContent
    }
}

# Step 6: Clean <body> tag attributes
$html = Strip-Attributes-InBody $html

# Step 7: Save cleaned HTML
$outputPath = "$env:TEMP\cleaned_output.html"
$html | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Cleaned HTML saved to: $outputPath"
