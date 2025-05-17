# Strip all <script>...</script> blocks from raw HTML
$decodedContent = [regex]::Replace(
    $decodedContent,
    '<script[^>]*>.*?</script\s*>',
    '',
    [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
)
