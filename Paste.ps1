$htmlBytes = Get-Content $htmlPath -Encoding Byte
$decodedContent = [System.Text.Encoding]::GetEncoding("windows-1252").GetString($htmlBytes)

# Nuke all comments
$decodedContent = [regex]::Replace(
    $decodedContent,
    '<!--[\s\S]*?-->',
    '',
    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
)

# Now load clean content
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($decodedContent)
