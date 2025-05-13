# Path to HtmlAgilityPack.dll
$hpackPath = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackPath

# Input/output paths
$inputPath = "C:\Path\To\Your\Input.html"
$outputPath = "C:\Path\To\Your\Output.html"

# Load HTML with Windows-1252 encoding
$encoding = [System.Text.Encoding]::GetEncoding("windows-1252")
$htmlContent = Get-Content -Path $inputPath -Encoding Byte
$decodedContent = $encoding.GetString($htmlContent)

# Load HTML
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($decodedContent)

# Create clean output doc
$cleanDoc = New-Object HtmlAgilityPack.HtmlDocument
$body = $cleanDoc.CreateElement("body")
$cleanDoc.DocumentNode.AppendChild($body)

# Bullet regexes and list type inference
$bulletTypes = @(
    @{ Regex = '^\s*(\d+)\.'; Type = '1' },
    @{ Regex = '^\s*([a-z])\.'; Type = 'a' },
    @{ Regex = '^\s*([A-Z])\.'; Type = 'A' },
    @{ Regex = '^\s*(i{1,3}|iv|v|vi{0,3}|ix|x)\.'; Type = 'i' }  # Lowercase roman
)

function Get-ListType ($text) {
    foreach ($entry in $bulletTypes) {
        if ($text -match $entry.Regex) {
            return $entry.Type
        }
    }
    return $null
}

$paragraphs = $doc.DocumentNode.SelectNodes("//p | //h1 | //h2 | //h3 | //table")
$currentList = $null
$currentListType = $null

foreach ($node in $paragraphs) {
    $text = $node.InnerText.Trim()

    if ([string]::IsNullOrWhiteSpace($text)) {
        continue
    }

    $detectedType = Get-ListType $text

    if ($detectedType) {
        # Start new list if type changes
        if ($null -eq $currentList -or $currentListType -ne $detectedType) {
            $currentList = $cleanDoc.CreateElement("ol")
            $currentList.SetAttributeValue("type", $detectedType)
            $body.AppendChild($currentList)
            $currentListType = $detectedType
        }

        # Remove bullet marker
        $cleanText = $text -replace '^\s*\S+\.\s*', ''

        $li = $cleanDoc.CreateElement("li")
        $li.InnerHtml = $cleanText
        $currentList.AppendChild($li)
    }
    else {
        # Close list if needed
        $currentList = $null
        $currentListType = $null

        # Preserve structural tags
        $tagName = $node.Name
        if ($tagName -in @("h1", "h2", "h3", "table")) {
            $copy = $cleanDoc.CreateElement($tagName)
            $copy.InnerHtml = $node.InnerHtml
            $body.AppendChild($copy)
        }
        else {
            $p = $cleanDoc.CreateElement("p")
            $p.InnerHtml = $node.InnerHtml
            $body.AppendChild($p)
        }
    }
}

# Save result
$cleanDoc.Save($outputPath)
Write-Host "Clean HTML saved to $outputPath"
