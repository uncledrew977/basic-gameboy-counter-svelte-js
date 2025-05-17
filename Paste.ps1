$headNode = $doc.DocumentNode.SelectSingleNode("//head")
if ($headNode -ne $null) {
    foreach ($child in @($headNode.ChildNodes)) {
        $text = $child.InnerHtml

        if ($text -match '<script[^>]*language\s*=\s*["'']?javascript["'']?[^>]*>.*?</script>' -or
            $text -match 'function\s+\w+\s*\(' -or
            $text -match '<!--\s*function') {

            $headNode.RemoveChild($child)
        }
    }
}
