# Create Word application
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Open the Word document
$docPath = "C:\Path\To\Your\Document.docx"
$doc = $word.Documents.Open($docPath)

# Define margin-left per list level (in points)
# Example: level 1 = 25pt, level 2 = 30pt, level 3 = 35pt, etc.
$indentMap = @{
    1 = 25
    2 = 30
    3 = 35
    4 = 40
    5 = 45
    6 = 50
    7 = 55
    8 = 60
    9 = 65
}

# Loop through all paragraphs
foreach ($para in $doc.Paragraphs) {
    $range = $para.Range
    $listFormat = $range.ListFormat

    if ($listFormat.ListType -ne 0) {
        $level = $listFormat.ListLevelNumber
        if ($indentMap.ContainsKey($level)) {
            $range.ParagraphFormat.LeftIndent = $indentMap[$level]
        } else {
            # Default fallback for levels not mapped
            $range.ParagraphFormat.LeftIndent = 70
        }
    }
}

# Save changes (overwrite the docx)
$doc.Save()

# OPTIONAL: Export to Filtered HTML
# $htmlOut = "C:\Path\To\Output.html"
# $doc.SaveAs([ref] $htmlOut, [ref] 10)  # 10 = wdFormatFilteredHTML

# Close and release
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
