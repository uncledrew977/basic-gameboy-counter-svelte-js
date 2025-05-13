# Path to HtmlAgilityPack.dll
$hpackPath = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackPath

# Input HTML path
$inputPath = "C:\Path\To\Your\Input.html"
$outputPath = "C:\Path\To\Your\Output.html"

# Load HTML with proper encoding (Windows-1252)
$encoding = [System.Text.Encoding]::GetEncoding("windows-1252")
$htmlContent = Get-Content -Path $inputPath -Encoding Byte
$decodedContent = $encoding.GetString($htmlContent)

# Load HTML into HtmlAgilityPack
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($decodedContent)

# Create a new HTML document for clean output
$cleanDoc = New-Object HtmlAgilityPack.HtmlDocument
$body = $cleanDoc.CreateElement("body")
$cleanDoc.DocumentNode.AppendChild($body)

# Regexes for bullet detection
$bulletRegexes = @(
    '^\s*(\d+)\.',         # 1., 2.
    '^\s*([a-zA-Z])\.',    # a., b., A., B.
    '^\s*(i{1,3}|iv|v|vi{0,3}|ix|x)\.' # i., ii., iii., iv., etc.
)

# Utility: Check if text is a bullet
function Is-Bullet ($text) {
    foreach ($regex in $bulletRegexes) {
        if ($text -match $regex) {
            return $true
        }
    }
    return $false
}

# Process nodes
$paragraphs = $doc.DocumentNode.SelectNodes("//p | //h1 | //h2 | //h3 | //table")

$listStack = @()
$currentList = $null

foreach ($node in $paragraphs) {
    $text = $node.InnerText.Trim()

    if ([string]::IsNullOrWhiteSpace($text)) {
        continue
    }

    if (Is-Bullet $text) {
        if (-not $currentList) {
            $currentList = $cleanDoc.CreateElement("ul")
            $body.AppendChild($currentList)
        }

        # Remove bullet marker
        $cleanText = $text -replace '^\s*\w+\.\s*', ''

        $li = $cleanDoc.CreateElement("li")
        $li.InnerHtml = $cleanText
        $currentList.AppendChild($li)
    }
    else {
        $currentList = $null

        $tagName = $node.Name
        if ($tagName -in @("h1", "h2", "h3", "table")) {
            $imported = $cleanDoc.CreateElement($tagName)
            $imported.InnerHtml = $node.InnerHtml
            $body.AppendChild($imported)
        }
        else {
            $p = $cleanDoc.CreateElement("p")
            $p.InnerHtml = $node.InnerHtml
            $body.AppendChild($p)
        }
    }
}

# Save cleaned HTML
$cleanDoc.Save($outputPath)
Write-Host "Clean HTML saved to $outputPath"
