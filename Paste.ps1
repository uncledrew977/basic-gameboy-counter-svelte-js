# Create Word application
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Open the Word document
$docPath = "C:\Path\To\Your\Document.docx"
$doc = $word.Documents.Open($docPath)

# Define indent per level in points
$indentMap = @{
    1 = 25
    2 = 40
    3 = 55
    4 = 70
    5 = 85
}

# Loop through paragraphs
foreach ($para in $doc.Paragraphs) {
    $range = $para.Range
    $listFormat = $range.ListFormat

    if ($listFormat.ListType -ne 0) {
        $level = $listFormat.ListLevelNumber

        if ($indentMap.ContainsKey($level)) {
            $indent = $indentMap[$level]

            # Apply paragraph indent
            $range.ParagraphFormat.LeftIndent = $indent
            $range.ParagraphFormat.FirstLineIndent = -18  # Optional: hanging indent

            # Force list level properties (Word often overrides paragraph indents without this)
            $template = $listFormat.ListTemplate
            if ($template -ne $null) {
                $listLevel = $template.ListLevels.Item($level)
                $listLevel.NumberFormat = $listLevel.NumberFormat  # Preserve format
                $listLevel.Alignment = 0  # Left
                $listLevel.NumberPosition = 0
                $listLevel.TextPosition = $indent
                $listLevel.TabPosition = $indent
            }
        }
    }
}

# Save the document
$doc.Save()

# Optional: Export to Filtered HTML
# $doc.SaveAs([ref] "C:\Path\To\Output.html", [ref] 10)

# Clean up
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
