$decodedContent = [regex]::Replace(
    $decodedContent,
    '<\s*script\b[^>]*>[\s\S]*?<\s*/\s*script\s*>',
    '',
    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
)
