# Paths
$htmlPath = "C:\Path\To\Test.html"
$outputPath = "$env:TEMP\fixed_lists_cleaned.html"

# Load HtmlAgilityPack DLL
$hpackDll = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackDll

# Load the HTML
$html = Get-Content $htmlPath -Raw -Encoding UTF8
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

# Setup
$indentUnit = 36
$listStack = @()
$outputBuilder = [ref] ''

function Close-Lists {
    param([int]$targetLevel, [ref]$builder)
    while ($listStack.Count -gt $targetLevel) {
        $builder.Value += "</$($listStack[-1])>`n"
        $listStack = $listStack[0..($listStack.Count - 2)]
    }
}

function Open-List {
    param($type, [ref]$builder)
    $builder.Value += "<$type>`n"
    $listStack += $type
}

function Is-Bullet($text) {
    return $text -match '^\s*(\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.\)\:]|\•|\▪|\·)\s+'
}

function Strip-Bullet($text) {
    return ($text -replace '^\s*(\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.\)\:]|\•|\▪|\·)\s+', '').Trim()
}

# Process nodes
foreach ($node in $doc.DocumentNode.SelectNodes("//p")) {
    $class = $node.GetAttributeValue("class", "")
    $style = $node.GetAttributeValue("style", "")
    $text = $node.InnerText.Trim()

    $level = 0
    if ($style -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
        $margin = [double]$matches[1]
        $level = [math]::Round($margin / $indentUnit)
    }

    if ($class -like "*MsoListParagraph*" -and (Is-Bullet $text)) {
        $cleanText = Strip-Bullet $text
        $listType = if ($text -match '^\s*\(?[0-9ivxlcdm]+\)?[\.\):]') { "ol" } else { "ul" }

        Close-Lists -targetLevel $level -builder $outputBuilder
        if ($listStack.Count -lt ($level + 1)) {
            Open-List -type $listType -builder $outputBuilder
        }

        $outputBuilder.Value += "<li>$cleanText</li>`n"
        $node.Remove()  # Remove this <p> from the original doc
    }
}

# Close any remaining lists
Close-Lists -targetLevel 0 -builder $outputBuilder

# Add remaining non-list content
$outputBuilder.Value += $doc.DocumentNode.InnerHtml

# Save result
Set-Content -Path $outputPath -Value $outputBuilder.Value -Encoding UTF8
Write-Output "Cleaned HTML with lists saved to: $outputPath"
