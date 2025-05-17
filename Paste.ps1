# Load raw HTML content
$htmlBytes = Get-Content $htmlPath -Encoding Byte
$decodedContent = [System.Text.Encoding]::GetEncoding("windows-1252").GetString($htmlBytes)

# 1. Remove only non-structural HTML comments (skip ones that contain <head>, <html>, or <body>)
$decodedContent = [regex]::Replace(
    $decodedContent,
    '<!--(?!.*?(<head|<html|<body)).*?-->',
    '',
    [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
)

# 2. Remove all <script>...</script> blocks (multiline, any case)
$decodedContent = [regex]::Replace(
    $decodedContent,
    '<\s*script[^>]*>.*?<\s*/\s*script\s*>',
    '',
    [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
)

# Now load into HtmlAgilityPack
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($decodedContent)
