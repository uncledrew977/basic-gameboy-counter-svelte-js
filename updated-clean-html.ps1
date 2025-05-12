# Load the HTML file
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Raw

# Step 1: Convert Word-generated lists into real HTML lists
function Convert-WordLists {
    param($htmlContent)

    $pattern = '<p[^>]*(MsoListParagraph|margin-left:[^">]+)[^>]*>(.*?)</p>'
    $matches = [regex]::Matches($htmlContent, $pattern, 'IgnoreCase')

    $output = ''
    $listStack = @()
    $prevLevel = 0

    foreach ($match in $matches) {
        $fullTag = $match.Value
        $innerHtml = $match.Groups[2].Value.Trim()

        if ($fullTag -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
            $margin = [double]($matches[1].Groups[1].Value)
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

    $htmlContent = $htmlContent -replace $pattern, ''
    return $htmlContent + $output
}

# Run list converter first
$html = Convert-WordLists $html

# Step 2: Remove class=""
$html = $html -replace '\sclass=""', ''

# Step 3: Remove style unless it contains background:silver
$html = [regex]::Replace($html, '\sstyle="([^"]*)"', {
    param($match)
    $style = $match.Groups[1].Value
    if ($style -match 'background\s*:\s*silver') {
        return $match.Value
    } else {
        return ''
    }
})

# Step 4: Strip attributes inside <body>, preserving background:silver
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

        return $htmlContent -replace [regex]::Escape($matches[0]), $bodyOpen + $cleanBodyContent + $bodyClose
    } else {
        return $htmlContent
    }
}

# Apply attribute cleaning
$html = Strip-Attributes-InBody $html

# Step 5: Save output with Set-Content (no emoji)
$outputPath = "$env:TEMP\cleaned_output.html"
$html | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Cleaned HTML saved to: $outputPath"
