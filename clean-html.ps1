# Load the HTML file
$htmlPath = "C:\Path\To\yourfile.html"
$html = Get-Content $htmlPath -Raw

# Remove class=""
$html = $html -replace '\sclass=""', ''

# Remove style="..." unless it contains 'background:silver'
$html = $html -replace '(\sstyle="(?![^"]*background\s*:\s*silver)[^"]*")', ''

# Function to strip attributes from tags inside <body>...</body>
function Strip-Attributes-InBody {
    param($htmlContent)

    # Match content inside <body>...</body>
    if ($htmlContent -match '(?s)(<body[^>]*>)(.*?)(</body>)') {
        $bodyOpen = $matches[1]
        $bodyContent = $matches[2]
        $bodyClose = $matches[3]

        # Strip all attributes from tags in $bodyContent
        $cleanBodyContent = $bodyContent -replace '<(\w+)(\s[^>]*?)?>', '<$1>'

        return $htmlContent -replace [regex]::Escape($matches[0]), "$bodyOpen$cleanBodyContent$bodyClose"
    } else {
        return $htmlContent
    }
}

# Strip all attributes from elements inside <body>
$html = Strip-Attributes-InBody $html

# Save the cleaned HTML
$html | Set-Content "C:\Path\To\cleaned_output.html"
