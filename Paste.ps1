function Remove-EmptyElementsRecursively {
    param([HtmlAgilityPack.HtmlNode]$rootNode)

    $removedAny = $false
    $nodes = $rootNode.SelectNodes(".//*")
    if ($nodes) {
        foreach ($node in $nodes | Sort-Object { $_.XPath.Length } -Descending) {
            # Skip any node that has attributes (like <link rel="..."> or <img src="...">)
            if ($node.Attributes.Count -gt 0) {
                continue
            }

            $text = $node.InnerText -replace '[\u00A0\u200B\s]+', '' -replace '&nbsp;', ''
            $hasMeaningfulChildren = $node.ChildNodes | Where-Object {
                $_.NodeType -ne 'Text' -and $_.Name -ne "#text"
            }

            if (-not $hasMeaningfulChildren -and [string]::IsNullOrWhiteSpace($text)) {
                $node.ParentNode.RemoveChild($node)
                $removedAny = $true
            }
        }
    }

    return $removedAny
}
