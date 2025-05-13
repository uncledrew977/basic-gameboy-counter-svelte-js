# Input and output paths
$htmlPath = "C:\Path\To\Test.html"
$outputPath = "$env:TEMP\fixed_lists_cleaned_debug.html"

# Load HtmlAgilityPack DLL
$hpackDll = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackDll

# Read HTML using correct encoding
$html = Get-Content $htmlPath -Raw -Encoding Default

# Wrap in body to ensure valid DOM
$html = "<html><body>" + $html + "</body></html>"

# Load HTML into HtmlAgilityPack
$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

# Count <p> tags safely
$paragraphs = $doc.DocumentNode.SelectNodes("//p")
if ($paragraphs) {
    Write-Output "Paragraph count: $($paragraphs.Count)"
} else {
    Write-Output "Paragraph count: 0 (no <p> elements found)"
}

# Constants
$indentUnit = 36
$listStack = @()
$outputBuilder = [ref] ''

# List handling helpers
function Close-Lists {
    param([int]$targetLevel, [ref]$builder)
    while ($listStack.Count -gt $targetLevel) {
        $builder.Value += "</$($listStack[-1])>`n"
        $listStack = $listStack[0..($listStack.Count - 2)]
    }
}

function Open-List {
    param($type, [ref]$builder)
    $builder.Value += "<$type>`n"
    $listStack += $type
}

# Bullet detection (non-breaking space safe)
function Is-Bullet {
    param($text)
    return $text -match '^[ \t\u00A0]*((\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.):]|[•▪·]))[ \t\u00A0]+'
}

function Strip-Bullet {
    param($text)
    return ($text -replace '^[ \t\u00A0]*((\(?[a-zA-Z0-9ivxlcdm]{1,5}[\.):]|[•▪·]))[ \t\u00A0]+', '').Trim()
}

# Process each paragraph
foreach ($node in $paragraphs) {

    # ✅ Skip paragraphs inside <table> cells
    if ($node.Ancestors("table").Count -gt 0) {
