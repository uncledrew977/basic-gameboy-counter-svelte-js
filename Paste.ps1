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

# Get bullet list template
$template = $word.ListGalleries.Item(1).ListTemplates.Item(1)  # 1 = bulleted list gallery

# Loop through all paragraphs
foreach ($para in $doc.Paragraphs) {
    $range = $para.Range
    $listFormat = $range.ListFormat

    if ($listFormat.ListType -ne 0) {
        $level = $listFormat.ListLevelNumber

        if ($indentMap.ContainsKey($level)) {
            $indent = $indentMap[$level]

            # Corrected method call without named arguments
            $listFormat.ApplyListTemplateWithLevel($template, $true, 1, $level)

            # Adjust template list level formatting
            $listLevel = $template.ListLevels.Item($level)
            $listLevel.NumberPosition = 0
            $listLevel.TextPosition = $indent
            $listLevel.TabPosition = $indent
            $listLevel.Alignment = 0

            # Apply paragraph indent and hanging indent
            $range.ParagraphFormat.LeftIndent = $indent
            $range.ParagraphFormat.FirstLineIndent = -18
        }
    }
}

# Save and cleanup
$doc.Save()
# Optional: $doc.SaveAs([ref] "C:\Path\To\Output.html", [ref] 10)

$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
