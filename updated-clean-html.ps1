$htmlPath = "C:\Path\To\Original.html"
$hpackDll = "C:\HtmlTools\HtmlAgilityPack\HtmlAgilityPack.1.11.46\lib\net45\HtmlAgilityPack.dll"
Add-Type -Path $hpackDll

$html = Get-Content $htmlPath -Raw -Encoding Default
$html = "<html><body>" + $html + "</body></html>"

$doc = New-Object HtmlAgilityPack.HtmlDocument
$doc.LoadHtml($html)

$nodes = $doc.DocumentNode.SelectNodes("//*[not(self::script or self::style)]/text()")
$i = 1
foreach ($textNode in $nodes) {
    $text = $textNode.InnerText
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $charCodes = $text.ToCharArray() | ForEach-Object { [int][char]$_ }
        Write-Output "`n[$i] TEXT: '$text'"
        Write-Output "CHAR CODES: $($charCodes -join ', ')"
        $i++
    }
}
