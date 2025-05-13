# Load HtmlAgilityPack
Add-Type -Path "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"

$inputPath = "C:\Test\test.html"
$outputPath = "C:\Test\output_cleaned.html"

# Load HTML with Windows-1252 encoding
$htmlRaw = Get-Content $inputPath -Encoding Default -Raw
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($htmlRaw)

# Remove all style and class attributes
$doc.DocumentNode.Descendants() | ForEach-Object {
    $_.Attributes.Remove("style")
    $_.Attributes.Remove("class")
}

# Bullet recognition
function IsBullet($text) {
    return $text -match '^\s*((\d+|[a-z]{1,3}|[ivxlc]{1,4})[\.\)]|•|·)\s+$'
}
function CleanBullet($text) {
    return $text -replace '^\s*((\d+|[a-z]{1,3}|[ivxlc]{1,4})[\.\)]|•|·)\s+', ''
}

# Detect and convert bullet paragraphs into real lists
$pNodes = $doc.DocumentNode.SelectNodes("//p")
if ($pNodes) {
    $listStack = @()
    $lastIndent = -1

    for ($i = 0; $i -lt $pNodes.Count; $i++) {
        $node = $pNodes[$i]
        $text = $node.InnerText.Trim()

        if ($text -match '^\s*((\d+|[a-z]{1,3}|[ivxlc]{1,4})[\.\)]|•|·)\s+') {
            $indent = ($node.InnerHtml -split '&nbsp;').Length - 1
            while ($listStack.Count -gt 0 -and $indent -le $listStack[-1].Indent) {
                $listStack[-1].ListNode = $listStack[-1].ListNode.ParentNode
                $listStack.Pop() | Out-Null
            }

            $listType = if ($text -match '^\s*\d+[\.\)]') { "ol" } else { "ul" }
            $li = $doc.CreateElement("li")
            $li.InnerHtml = CleanBullet $node.InnerHtml

            if ($listStack.Count -eq 0 -or $indent -gt $lastIndent) {
                $newList = $doc.CreateElement($listType)
                $node.ParentNode.InsertBefore($newList, $node)
                $listStack.Push([PSCustomObject]@{ ListNode = $newList; Indent = $indent })
            }

            $listStack[-1].ListNode.AppendChild($li)
            $node.ParentNode.RemoveChild($node)
            $lastIndent = $indent
        }
    }
}

# Save output HTML
$doc.Save($outputPath)
Write-Host "Cleaned HTML saved to: $outputPath"
