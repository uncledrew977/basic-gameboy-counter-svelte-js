function Remove-EmptyElementsRecursively {
    param([HtmlAgilityPack.HtmlNode]$rootNode)

    $removedAny = $false
    $nodes = $rootNode.SelectNodes(".//*")
    if ($nodes) {
        # Process from bottom-up so we handle child-empty-parent chains
        foreach ($node in $nodes | Sort-Object { $_.XPath.Length } -Descending) {
            if ($node.Name -in @("img", "br", "hr")) { continue }

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

# Run the cleanup until nothing else is removed (recursive upward pass)
do {
    $more = Remove-EmptyElementsRecursively -rootNode $doc.DocumentNode
} while ($more)
