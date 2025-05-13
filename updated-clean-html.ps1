# Input and output paths
$htmlPath = "C:\Path\To\Test.html"
$outputPath = "$env:TEMP\fixed_lists_cleaned.html"

# Load HtmlAgilityPack DLL (adjust this path if needed)
$hpackDll = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackDll

# Load HTML content
$html = Get-Content $htmlPath -Raw -Encoding UTF8
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

# Constants and list tracking
$indentUnit = 36
$listStack = @()
$outputBuilder = [ref] ''

# Helper to close open lists to the given level
function Close-Lists {
    param([int]$targetLevel, [ref]$builder)
    while ($listStack.Count -gt $targetLevel) {
        $builder.Value += "</$($listStack[-1])>`n"
        $listStack = $listStack[0..($listStack.Count - 2)]
    }
}

# Helper to open a new list
function Open-List {
    param($type, [ref]$builder)
    $builder.Value += "<$type>`n"
    $listStack += $type
}

# Detects if the line starts with a bullet
function Is-Bullet {
    param($text)
    return $text -match '^\s*((\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.):]|[•▪·]))\s+'
}

# Removes the bullet prefix
function Strip-Bullet {
    param($text)
    return ($text -replace '^\s*((\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.):]|[•▪·]))\s+', '').Trim()
}

# Process all <p> tags — most Word list items live here
foreach ($node in $doc.DocumentNode.SelectNodes("//p")) {
    $class = $node.GetAttributeValue("class", "")
    $style = $node.GetAttributeValue("style", "")
    $text = $node.InnerText.Trim()

    # Determine indent level
    $level = 0
    if ($style -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
        $margin = [double]$matches[1]
        $level = [math]::Round($margin / $indentUnit)
    }

    # Check if it's a Word-style list item
    if ($class -like "*MsoListParagraph*" -and (Is-Bullet $text)) {
        $cleanText = Strip-Bullet $text
        $listType = if ($text -match '^\s*\(?[0-9ivxlcdm]+\)?[\.):]') { "ol" } else { "ul" }

        # Open/close lists as needed
        Close-Lists -targetLevel $level -builder $outputBuilder
        if ($listStack.Count -lt ($level + 1)) {
            Open-List -type $listType -builder $outputBuilder
        }

        $outputBuilder.Value += "<li>$cleanText</li>`n"
        $node.Remove()  # Don't output the original <p>
    }
}

# Close any remaining open lists
Close-Lists -targetLevel 0 -builder $outputBuilder

# Add back remaining (non-list) HTML content
$outputBuilder.Value += $doc.DocumentNode.InnerHtml

# Save result
Set-Content -Path $outputPath -Value $outputBuilder.Value -Encoding UTF8
Write-Output "✅ Cleaned HTML saved to: $outputPath"
