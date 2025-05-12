param (
    [string]$InputPath,
    [string]$OutputPath
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($InputPath, $false, $true)
    $doc.SaveAs($OutputPath, 10)  # 10 = wdFormatFilteredHTML
    $doc.Close()
    Write-Host "Converted: $InputPath"
} catch {
    Write-Warning "Failed: $InputPath - $_"
}

$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
