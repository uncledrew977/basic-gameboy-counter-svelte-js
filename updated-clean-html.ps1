
# Load HtmlAgilityPack
Add-Type -Path "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.dll"

$inputPath = "C:\Path\To\Input\test (3).html"
$outputPath = "C:\Path\To\Output\cleaned_test3.html"

$html = Get-Content -Path $inputPath -Encoding Default -Raw
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

function Get-ListType {
    param($text)

    if ($text -match '^\(?[ivxlc]+\)\.?$') { return "ol" }    # roman numerals
    elseif ($text -match '^[a-zA-Z]\.$') { return "ol" }      # alphabetic
    elseif ($text -match '^\d+\.$') { return "ol" }           # numeric
    elseif ($text -match '^[•·▪]$') { return "ul" }           # bullet chars
    else { return $null }
}

function CleanNode {
    param([HtmlAgilityPack.HtmlNode]$node, [HtmlAgilityPack.HtmlNode]$parent)

    switch ($node.Name.ToLower()) {
        "h1" { $newNode = $node.Clone(); $parent.AppendChild($newNode) }
        "h2" { $newNode = $node.Clone(); $parent.AppendChild($newNode) }
        "h3" { $newNode = $node.Clone(); $parent.AppendChild($newNode) }
        "h4" { $newNode = $node.Clone(); $parent.AppendChild($newNode) }
        "h5" { $newNode = $node.Clone(); $parent.AppendChild($newNode) }
        "h6" { $newNode = $node.Clone(); $parent.AppendChild($newNode) }
        "p" {
            $text = $node.InnerText.Trim()
            if ($text -match '^(?<bullet>(\(?[ivxlc]+\)|[a-zA-Z]|\d+|[•·▪]))[\.\)]\s+(?<content>.+)') {
                $bullet = $matches['bullet']
                $content = $matches['content']
                $listType = Get-ListType $bullet
                if ($listType) {
                    if (-not $script:currentList -or $script:currentList.Name -ne $listType) {
                        $script:currentList = $doc.CreateElement($listType)
                        $parent.AppendChild($script:currentList)
                    }
                    $li = $doc.CreateElement("li")
                    $li.InnerHtml = $content
                    $script:currentList.AppendChild($li)
                    return
                }
            }
            $script:currentList = $null
            $newP = $doc.CreateElement("p")
            $newP.InnerHtml = $node.InnerHtml
            $parent.AppendChild($newP)
        }
        "table" {
            $newTable = $node.Clone()
            $parent.AppendChild($newTable)
        }
        default {
            foreach ($child in $node.ChildNodes) {
                CleanNode -node $child -parent $parent
            }
        }
    }
}

$newDoc = New-Object HtmlAgilityPack.HtmlDocument
$htmlElem = $newDoc.CreateElement("html")
$bodyElem = $newDoc.CreateElement("body")
$newDoc.DocumentNode.AppendChild($htmlElem)
$htmlElem.AppendChild($bodyElem)

$script:currentList = $null
CleanNode -node $doc.DocumentNode.DocumentNode -parent $bodyElem

# Save output
$newDoc.Save($outputPath)
Write-Host "Cleaned HTML saved to: $outputPath"
