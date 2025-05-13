
# Load HtmlAgilityPack
Add-Type -Path "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.dll"

$inputPath = "C:\Path\To\Your\test (3).html"
$outputPath = "C:\Path\To\Your\cleaned_output.html"

$html = Get-Content -Path $inputPath -Raw -Encoding Default
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

$bulletPatterns = @(
    '^\s*(\d+)\.\s*',       # 1., 2., etc.
    '^\s*([a-zA-Z])\.\s*',  # a., b., A., B., etc.
    '^\s*(i{1,3}|iv|v|vi{1,3}|ix|x)\.\s*', # i., ii., iii., iv., etc.
    '^\s*[\u2022\u00B7·]\s*'  # Unicode bullets: • ·
)

function IsBullet($text) {
    foreach ($pattern in $bulletPatterns) {
        if ($text -match $pattern) { return $true }
    }
    return $false
}

function StripBullet($text) {
    foreach ($pattern in $bulletPatterns) {
        if ($text -match $pattern) {
            return $text -replace $pattern, ''
        }
    }
    return $text
}

function CreateCleanNode($node) {
    switch ($node.Name.ToLower()) {
        "h1" { return $node.Clone() }
        "h2" { return $node.Clone() }
        "h3" { return $node.Clone() }
        "h4" { return $node.Clone() }
        "h5" { return $node.Clone() }
        "h6" { return $node.Clone() }
        "table" { return $node.Clone() }
        "thead" { return $node.Clone() }
        "tbody" { return $node.Clone() }
        "tr" { return $node.Clone() }
        "th" { return $node.Clone() }
        "td" { return $node.Clone() }
        "p" {
            $text = $node.InnerText.Trim()
            if (IsBullet($text)) {
                return @{ type = "li"; content = StripBullet($text); raw = $node }
            } elseif ($text) {
                $p = $doc.CreateElement("p")
                $p.InnerHtml = $node.InnerHtml
                return $p
            }
        }
        default { return $null }
    }
}

$body = $doc.DocumentNode.SelectSingleNode("//body")
$newBody = $doc.CreateElement("body")
$currentList = $null

foreach ($child in $body.ChildNodes) {
    $clean = CreateCleanNode $child
    if ($clean -is [HtmlAgilityPack.HtmlNode]) {
        if ($currentList) {
            $newBody.AppendChild($currentList)
            $currentList = $null
        }
        $newBody.AppendChild($clean)
    } elseif ($clean -is [Hashtable] -and $clean.type -eq "li") {
        if (-not $currentList) {
            $currentList = $doc.CreateElement("ul")
        }
        $li = $doc.CreateElement("li")
        $li.InnerHtml = $clean.content
        $currentList.AppendChild($li)
    }
}

if ($currentList) {
    $newBody.AppendChild($currentList)
}

$doc.DocumentNode.SelectSingleNode("//body").RemoveAllChildren()
$doc.DocumentNode.SelectSingleNode("//body").AppendChild($newBody)
$doc.Save($outputPath)
