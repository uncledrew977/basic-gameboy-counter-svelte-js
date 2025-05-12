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

# 3. Function to strip all attributes from elements inside <body>...</body>
function Strip-Attributes-InBody {
    param($htmlContent)

    if ($htmlContent -match '(?s)(<body[^>]*>)(.*?)(</body>)') {
        $bodyOpen = $matches[1]
        $bodyContent = $matches[2]
        $bodyClose = $matches[3]

        # Remove all attributes from tags inside the body
        $cleanBodyContent = $bodyContent -replace '<(\w+)(\s[^>]*?)?>', '<$1>'

        return $htmlContent -replace [regex]::Escape($matches[0]), "$bodyOpen$cleanBodyContent$bodyClose"
    } else {
        return $htmlContent
    }
}

# Apply body attribute stripping
$html = Strip-Attributes-InBody $html

# Save the cleaned HTML using .NET (for restricted environments)
$outputPath = "$env:TEMP\cleaned_output.html"
[System.IO.File]::WriteAllText($outputPath, $html)
Write-Output "✅ Cleaned HTML saved to: $outputPath"
