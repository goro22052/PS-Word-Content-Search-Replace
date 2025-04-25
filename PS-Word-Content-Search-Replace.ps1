
$startTime = Get-Date

$times = 100
$docPath = "\\kyv-s-f02\FileShare\Папка обміну\IT\Насіння Кредит Застава Гривня ТАР з 10 05 2023.docx"
$oldWord = "Вступ"
$newWord = "!!!!!!!!"

$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false

$document = $wordApp.Documents.Open($docPath)

$selection = $wordApp.Selection

for ($i = 1; $i -le $times; $i++) {
    $range = $document.Content
    $find = $range.Find
    
    $find.Text = $oldWord
    $find.Replacement.Text = $newWord
    $find.Forward = $true
    $find.Wrap = 1            # wdFindContinue
    $find.Format = $false
    $find.MatchCase = $false
    $find.MatchWholeWord = $false

    $find.Execute(
        [ref]$oldWord,
        [ref]$false,
        [ref]$false,
        [ref]$false,
        [ref]$false,
        [ref]$false,
        [ref]$true,
        [ref]1,
        [ref]$false,
        [ref]$newWord,
        [ref]2,   # wdReplaceAll
        [ref]$false,
        [ref]$false,
        [ref]$false,
        [ref]$false
    ) | Out-Null
    }

$document.Save()
$document.Close()
$wordApp.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($find) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($selection) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null

[GC]::Collect()
[GC]::WaitForPendingFinalizers()

$endTime = Get-Date
$ElapsedTime = ($endTime - $startTime).TotalSeconds

Write-Host  "Час виконання $times ітерацій пошуку -  $ElapsedTime "  -ForegroundColor Yellow


