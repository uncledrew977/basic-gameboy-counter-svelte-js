# Input and output paths
$htmlPath = "C:\Path\To\Test.html"
$outputPath = "$env:TEMP\fixed_lists_cleaned.html"

# Load HtmlAgilityPack DLL (adjust path to match your version)
$hpackDll = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackDll

# Load the HTML content
$html = Get-Content $htmlPath -Raw -Encoding UTF8
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

# Constants and tracking
$indentUnit = 36
$listStack = @()
$outputBuilder = [ref] ''

# Close open lists down to target level
function Close-Lists {
    param([int]$targetLevel, [ref]$builder)
    while ($listStack.Count -gt $targetLevel) {
        $builder.Value += "</$($listStack[-1])>`n"
        $listStack = $listStack[0..($listStack.Count - 2)]
    }
}

# Open a new list
function Open-List {
    param($type, [ref]$builder)
    $builder.Value += "<$type>`n"
    $listStack += $type
}

# Match Word-style bullets (escaped properly for PowerShell/regex)
function Is-Bullet {
    param($text)
    return $text -match '^\s*((\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.\)\:]|[\•\▪\·]))\s+'
}

# Strip bullet prefix from text
function Strip-Bullet {
    param($text)
    return ($text -replace '^\s*((\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.\)\:]|[\•\▪\·]))\s+', '').Trim()
}

# Process <p> nodes (Word list items typically live here)
foreach ($node in $doc.DocumentNode.SelectNodes("//p")) {
    $class = $node.GetAttributeValue("class", "")
    $style = $node.GetAttributeValue("style", "")
    $text = $node.InnerText.Trim()

    # Determine nesting level from margin-left
    $level = 0
    if ($style -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
        $margin = [double]$matches[1]
        $level = [math]::Round($margin / $indentUnit)
    }

    # If it's a list paragraph and starts with a bullet...
    if ($class -like "*MsoListParagraph*" -and (Is-Bullet $text)) {
        $cleanText = Strip-Bullet $text
        $listType = if ($text -match '^\s*\(?[0-9ivxlcdm]+\)?[\.\):]') { "ol" } else { "ul" }

        # Open/close lists as needed
        Close-Lists -targetLevel $level -builder $outputBuilder
        if ($listStack.Count -lt ($level + 1)) {
            Open-List -type $listType -builder $outputBuilder
        }

        $outputBuilder.Value += "<li>$cleanText</li>`n"
        $node.Remove()  # Don't output this node again later
    }
}

# Close remaining open lists
Close-Lists -targetLevel 0 -builder $outputBuilder

# Append remaining non-list HTML
$outputBuilder.Value += $doc.DocumentNode.InnerHtml

# Save final output
Set-Content -Path $outputPath -Value $outputBuilder.Value -Encoding UTF8
Write-Output "✅ Cleaned HTML with lists saved to: $outputPath"
