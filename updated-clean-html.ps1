# Input and output paths
$htmlPath = "C:\Path\To\Test.html"
$outputPath = "$env:TEMP\fixed_lists_cleaned_debug.html"

# Load HtmlAgilityPack DLL (adjust if needed)
$hpackDll = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackDll

# Read HTML (with correct encoding)
$html = Get-Content $htmlPath -Raw -Encoding Default

# Wrap HTML to ensure valid DOM structure
$html = "<html><body>" + $html + "</body></html>"

# Load with HtmlAgilityPack
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

# Debug: how many paragraphs were found?
Write-Output "Paragraph count: $($doc.DocumentNode.SelectNodes('//p')?.Count)"

# Constants
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

function Is-Bullet {
    param($text)
    return $text -match '^\s*((\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.):]|[•▪·]))\s+'
}

function Strip-Bullet {
    param($text)
    return ($text -replace '^\s*((\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.):]|[•▪·]))\s+', '').Trim()
}

# Loop through all <p> nodes
foreach ($node in $doc.DocumentNode.SelectNodes("//p")) {
    $class = $node.GetAttributeValue("class", "")
    $style = $node.GetAttributeValue("style", "")
    $text = $node.InnerText.Trim()

    # Debug output for each paragraph
    Write-Output "`n---"
    Write-Output "CLASS: $class"
    Write-Output "TEXT: $text"
    Write-Output "IS BULLET: $(Is-Bullet $text)"

    # Estimate nesting level
    $level = 0
    if ($style -match 'margin-left:\s*(\d+(\.\d+)?)pt') {
        $margin = [double]$matches[1]
        $level = [math]::Round($margin / $indentUnit)
    }

    if ($class -like "*MsoListParagraph*" -and (Is-Bullet $text)) {
        $cleanText = Strip-Bullet $text
        $listType = if ($text -match '^\s*\(?[0-9ivxlcdm]+\)?[\.):]') { "ol" } else { "ul" }

        Close-Lists -targetLevel $level -builder $outputBuilder
        if ($listStack.Count -lt ($level + 1)) {
            Open-List -type $listType -builder $outputBuilder
        }

        $outputBuilder.Value += "<li>$cleanText</li>`n"
        $node.Remove()
    }
}

# Close remaining lists
Close-Lists -targetLevel 0 -builder $outputBuilder

# Add remaining HTML
$outputBuilder.Value += $doc.DocumentNode.InnerHtml

# Save output
Set-Content -Path $outputPath -Value $outputBuilder.Value -Encoding UTF8
Write-Output "✅ Cleaned HTML with debug saved to: $outputPath"
