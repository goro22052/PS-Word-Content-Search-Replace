$totalTimes = 10
$times = 30
$docPath = "C:\Users\Holubenko.Roman\Desktop\Обновление Zabbix 6.0 до 7.0_.docx"
$oldWord = "Вступ"
$newWord = "!!!!!!!!"
$SummaryTime = 0

Write-Host "Тестування Word.Content.Find. $totalTimes тестів по $times ітерацій. На $env:COMPUTERNAME." -ForegroundColor Cyan

for ($j = 1; $j -le $totalTimes; $j++) {
    $dt = Get-Date
    Write-Host "############# $dt Запуск тесту номер $j. ##########" -ForegroundColor Green
    $startTime = Get-Date

    try {
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $false

        for ($i = 1; $i -le $times; $i++) {
            $document = $wordApp.Documents.Open($docPath)

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

            $document.Save()
            $document.Close()
            
            # Звільнення COM-об’єктів документа після кожної ітерації
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($find) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
        }

        $wordApp.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null

    } catch {
        Write-Host "Сталася помилка: $_" -ForegroundColor Red
    } finally {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    $endTime = Get-Date
    $elapsedTime = ($endTime - $startTime).TotalSeconds
    Write-Host "Час виконання $times ітерацій пошуку – $elapsedTime секунд." -ForegroundColor Yellow
    Write-Host "############# $dt Тест номер $j закінчив виконання. ##########" -ForegroundColor Green
    $SummaryTime += $elapsedTime
    $pause = Get-Random -Minimum 30 -Maximum 120
    
    if ($j -lt $totalTimes) {
        $pause = Get-Random -Minimum 30 -Maximum 120
        Write-Host "Роблю паузу в $pause секунд." -ForegroundColor Gray
        Start-Sleep -Seconds $pause
    }

}

$averageTime = $SummaryTime / $totalTimes
Write-Host "Всі тести завершено. Середній час виконання $times ітерацій пошуку по документу - $averageTime секунд" -ForegroundColor Cyan
