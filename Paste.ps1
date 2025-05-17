# Launch Word
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Open DOCX
$docPath = "C:\Path\To\Your\Document.docx"
$doc = $word.Documents.Open($docPath)

# Bullet indents (points) per level
$indentMap = @{
    1 = 25
    2 = 50
    3 = 75
    4 = 100
    5 = 125
}

# Create a fresh bullet list template
$template = $word.ListGalleries.Item(1).ListTemplates.Item(1)  # 1 = bulleted lists

# Loop through all paragraphs
foreach ($para in $doc.Paragraphs) {
    $range = $para.Range
    $listFormat = $range.ListFormat

    # Check if it's a list item
    if ($listFormat.ListType -ne 0) {
        $level = $listFormat.ListLevelNumber

        if ($indentMap.ContainsKey($level)) {
            $indent = $indentMap[$level]

            # Force the list to reapply using our bullet template
            $listFormat.ApplyListTemplateWithLevel(
                $template,
                $ContinuePreviousList = $true,
                $DefaultListBehavior = 1,
                $ApplyLevel = $level
            )

            # Force list level indent properties
            $listLevel = $template.ListLevels.Item($level)
            $listLevel.NumberPosition = 0
            $listLevel.TextPosition = $indent
            $listLevel.TabPosition = $indent
            $listLevel.Alignment = 0  # Left aligned

            # Apply paragraph indent and optional hanging indent
            $range.ParagraphFormat.LeftIndent = $indent
            $range.ParagraphFormat.FirstLineIndent = -18
        }
    }
}

# Save updated doc
$doc.Save()

# OPTIONAL: Export to filtered HTML (shows margin-left)
# $doc.SaveAs([ref] "C:\Path\To\Output.html", [ref] 10)  # 10 = wdFormatFilteredHTML

# Clean up
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
