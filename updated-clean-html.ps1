# Load HtmlAgilityPack
Add-Type -Path "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.dll"

# Input and output paths
$inputPath = "C:\Test\test (3).html"
$outputPath = "C:\Test\clean_output.html"

# Load the document using the correct encoding
$encoding = [System.Text.Encoding]::GetEncoding("windows-1252")
$htmlContent = [System.IO.File]::ReadAllText($inputPath, $encoding)
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.OptionFixNestedTags = $true
$doc.LoadHtml($htmlContent)

# Clean attributes from a node
function Clean-Attributes($node) {
    $attrsToKeep = @("colspan", "rowspan") # Keep essential table attrs
    foreach ($attr in $node.Attributes.ToArray()) {
        if ($attrsToKeep -notcontains $attr.Name.ToLower()) {
            $node.Attributes.Remove($attr)
        }
    }
}

# Recursively clean attributes and remove non-structural tags
function Clean-Node($node) {
    $preserveTags = @("h1","h2","h3","h4","h5","h6","p","ul","ol","li","table","thead","tbody","tr","td","th")
    if ($node.NodeType -eq "Element") {
        if ($preserveTags -notcontains $node.Name.ToLower()) {
            $node.Name = "p"  # Replace unknown elements with <p>
        }
        Clean-Attributes $node
    }

    foreach ($child in $node.ChildNodes.ToArray()) {
        Clean-Node $child
    }
}

Clean-Node $doc.DocumentNode

# Save the cleaned HTML
$sw = New-Object System.IO.StreamWriter($outputPath, $false, $encoding)
$doc.Save($sw)
$sw.Close()

Write-Host "Cleaned HTML saved to $outputPath"
