# Load HtmlAgilityPack
Add-Type -Path "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.dll"

$inputPath = "C:\Test\input.html"
$outputPath = "C:\Test\cleaned_output.html"

# Load HTML content
$htmlRaw = Get-Content $inputPath -Encoding Default -Raw
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($htmlRaw)

# Detect list bullet patterns
function IsBullet($text) {
    return $text -match '^\s*((\d+|[a-z]{1,3}|[ivxlc]{1,4})[\.\)]|•|·)\s+'
}

function CleanBullet($html) {
    return $html -replace '^\s*((\d+|[a-z]{1,3}|[ivxlc]{1,4})[\.\)]|•|·)\s+', ''
}

function GetIndentLevel($html) {
    return ($html -split '&nbsp;').Count - 1
}

# Process all <p> elements
$pNodes = $doc.DocumentNode.SelectNodes("//p")
if ($pNodes) {
    $listStack = @()
    $lastIndent = -1

    foreach ($p in $pNodes) {
        $text = $p.InnerText.Trim()
        $html = $p.InnerHtml

        if (IsBullet $text) {
            $indent = GetIndentLevel $html
            $listType = if ($text -match '^\s*\d+[\.\)]') { "ol" } else { "ul" }

            while ($listStack.Count -gt 0 -and $indent -le $listStack[-1].Indent) {
                $listStack[-1].ListNode = $listStack[-1].ListNode.ParentNode
                $listStack.Pop() | Out-Null
            }

            if ($listStack.Count -eq 0 -or $indent -gt $lastIndent) {
                $newList = $doc.CreateElement($listType)
                $p.ParentNode.InsertBefore($newList, $p)
                $listStack.Push([PSCustomObject]@{ ListNode = $newList; Indent = $indent })
            }

            $li = $doc.CreateElement("li")
            $li.InnerHtml = CleanBullet $html
            $listStack[-1].ListNode.AppendChild($li)
            $p.Remove()
            $lastIndent = $indent
        }
    }
}

# Remove style/class attributes AFTER bullet conversion
$doc.DocumentNode.Descendants() | ForEach-Object {
    $_.Attributes.Remove("style")
    $_.Attributes.Remove("class")
}

# Save the cleaned HTML
$doc.Save($outputPath)
Write-Host "Clean HTML saved to $outputPath"
