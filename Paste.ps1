$decodedContent = [regex]::Replace(
    $decodedContent,
    '<script[^>]*language\s*=\s*["'']?javascript["'']?[^>]*>.*?</script>',
    '',
    [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
)
