$inputDir = "C:\path\to\docs"
$outputDir = "C:\path\to\html"
$wdFormatFilteredHTML = 10

if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false

Get-ChildItem -Path $inputDir -Filter *.docx | ForEach-Object {
    $doc = $word.Documents.Open($_.FullName)
    $outputPath = Join-Path $outputDir ($_.BaseName + ".html")
    $doc.SaveAs([ref] $outputPath, [ref] $wdFormatFilteredHTML)
    $doc.Close()
    Write-Host "Exported: $outputPath"

    # Load HTML
    $html = Get-Content $outputPath -Raw

    # --- Basic Cleanup ---
    $html = $html -replace '<!--\[if gte mso .*?\]>', ''
    $html = $html -replace '<meta name=Generator.*?>', ''
    $html = $html -replace '<!--.*?-->', ''
    $html = $html -replace 'xmlns:[^=]*="[^"]*"', ''
    $html = $html -replace 'class=""', ''
    $html = $html -replace '\s{2,}', ' '
    $html = $html -replace '<p>\s*</p>', ''

    # --- Selectively Keep Silver Background Style ---
    $html = $html -replace 'style="([^"]*)"' , {
        $style = $_.Groups[1].Value
        if ($style -match 'background:\s*silver') {
            'style="background:silver"'
        } else {
            ''
        }
    }

    # --- List Handling (Nested + Lettered) ---
    $htmlLines = $html -split "`r?`n"
    $listStack = @()
    $processed = @()

    foreach ($line in $htmlLines) {
        $isListItem = $false
        $listType = ""
        $itemContent = ""
        $indentLevel = 0
        $customListAttr = ""

        # Determine indentation by margin-left
        if ($line -match 'margin-left:(\d+)pt') {
            $indentLevel = [math]::Floor(($matches[1] -as [int]) / 18)
        }

        # Detect bulleted list
        if ($line -match '<p[^>]*?>\s*[•·]\s*(.*?)</p>') {
            $itemContent = $matches[1]
            $listType = "ul"
            $isListItem = $true
        }
        # Detect numbered list (1. or 1))
        elseif ($line -match '<p[^>]*?>\s*\d+[\.\)]\s*(.*?)</p>') {
            $itemContent = $matches[1]
            $listType = "ol"
            $isListItem = $true
        }
        # Detect lettered list (a. or a))
        elseif ($line -match '<p[^>]*?>\s*[a-zA-Z][\.\)]\s*(.*?)</p>') {
            $itemContent = $matches[1]
            $listType = "ol"
            $customListAttr = ' type="a"'
            $isListItem = $true
        }

        if ($isListItem) {
            # Adjust nesting level
            while ($listStack.Count -gt $indentLevel) {
                $processed += "</$($listStack[-1].Type)>"
                $listStack = $listStack[0..($listStack.Count - 2)]
            }
            while ($listStack.Count -lt $indentLevel) {
                $processed += "<$listType$customListAttr>"
                $listStack += @{ Type = $listType; Attr = $customListAttr }
            }

            if ($listStack.Count -eq 0 -or $listStack[-1].Type -ne $listType -or $listStack[-1].Attr -ne $customListAttr) {
                if ($listStack.Count -gt 0) {
                    $processed += "</$($listStack[-1].Type)>"
                    $listStack = $listStack[0..($listStack.Count - 2)]
                }
                $processed += "<$listType$customListAttr>"
                $listStack += @{ Type = $listType; Attr = $customListAttr }
            }

            $processed += "  <li>$itemContent</li>"
        }
        else {
            # Close all open lists if we're out of list context
            while ($listStack.Count -gt 0) {
                $processed += "</$($listStack[-1].Type)>"
                $listStack = $listStack[0..($listStack.Count - 2)]
            }
            $processed += $line
        }
    }

    # Final close
    while ($listStack.Count -gt 0) {
        $processed += "</$($listStack[-1].Type)>"
        $listStack = $listStack[0..($listStack.Count - 2)]
    }

    # Rebuild HTML
    $html = $processed -join "`r`n"
    Set-Content -Path $outputPath -Value $html
    Write-Host "Cleaned + Nested + Lettered Lists: $outputPath"
}

$word.Quit()
