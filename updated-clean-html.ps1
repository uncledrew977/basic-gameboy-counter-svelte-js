# Define paths
$wordFilePath = "C:\Path\To\Your\Document.docx"
$outputHtmlPath = "C:\Path\To\Output\Document.html"

# Create Word application object
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false

# Open the Word document
$document = $wordApp.Documents.Open($wordFilePath)

# Save as filtered HTML
$document.SaveAs([ref] $outputHtmlPath, [ref] 10)  # 10 corresponds to wdFormatFilteredHTML

# Close the document and quit Word
$document.Close()
$wordApp.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Output "Filtered HTML saved to: $outputHtmlPath"
