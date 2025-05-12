# Load the HTML file
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Raw

# 1. Remove class=""
$html = $html -replace '\sclass=""', ''

# 2. Remove style="" unless it contains background:silver
$html = [regex]::Replace($html, '\sstyle="([^"]*)"', {
    param($match)
    $style = $match.Groups[1].Value
    if ($style -match 'background\s*:\s*silver') {
        return $match.Value  # Keep the original style
    } else {
        return ''            # Remove the style attribute
    }
})

# 3. Function to strip attributes inside <body>, preserving only style with background:silver
function Strip-Attributes-InBody {
    param($htmlContent)

    if ($htmlContent -match '(?s)(<body[^>]*>)(.*?)(</body>)') {
        $bodyOpen    = $matches[1]
        $bodyContent = $matches[2]
        $bodyClose   = $matches[3]

        $cleanBodyContent = [regex]::Replace($bodyContent, '<(\w+)([^>]*)>', {
            param($tagMatch)
            $tag   = $tagMatch.Groups[1].Value
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

# 4. Apply body-cleaning
$html = Strip-Attributes-InBody $html

# 5. Save using Set-Content
$outputPath = "$env:TEMP\cleaned_output.html"
$html | Set-Content -Path $outputPath -Encoding UTF8
Write-Output "Cleaned HTML saved to: $outputPath"
