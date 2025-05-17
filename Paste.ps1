# Start Word COM application
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Path to your Word document
$docPath = "C:\Path\To\Your\Document.docx"
$doc = $word.Documents.Open($docPath)

# Define desired margin-left (in points) per bullet/number level
$indentMap = @{
    1 = 25
    2 = 50
    3 = 75
    4 = 100
    5 = 125
}

# Safely access the bulleted list gallery (2 = bulleted)
$galleryIndex = 2
$templateIndex = 1

$listGallery = $word.ListGalleries.Item($galleryIndex)

if ($null -eq $listGallery -or $listGallery.ListTemplates.Count -lt $templateIndex) {
    Write-Error "Could not access bulleted list template from ListGalleries.Item($galleryIndex)."
    $doc.Close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    exit
}

$template = $listGallery.ListTemplates.Item($templateIndex)

# Loop through all paragraphs in the document
foreach ($para in $doc.Paragraphs) {
    $range = $para.Range
    $listFormat = $range.ListFormat

    # Process only list items (bullets, numbers, letters, roman numerals)
    if ($listFormat.ListType -ne 0) {
        $level = $listFormat.ListLevelNumber

        if ($indentMap.ContainsKey($level)) {
            $indent = $indentMap[$level]

            # Apply the list template with this level
            $listFormat.ApplyListTemplateWithLevel($template, $true, 1, $level)

            # Force formatting for the specific list level
            $listLevel = $template.ListLevels.Item($level)
            $listLevel.NumberPosition = 0
            $listLevel.TextPosition = $indent
            $listLevel.TabPosition = $indent
            $listLevel.Alignment = 0  # Left aligned

            # Apply paragraph indentation (will be respected in HTML output)
            $range.ParagraphFormat.LeftIndent = $indent
            $range.ParagraphFormat.FirstLineIndent = -18  # Hanging indent
        }
    }
}

# Save the updated Word document
$doc.Save()

# Optional: Export to Filtered HTML
# $htmlPath = "C:\Path\To\Output.html"
# $doc.SaveAs([ref] $htmlPath, [ref] 10)  # 10 = wdFormatFilteredHTML

# Clean up
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
