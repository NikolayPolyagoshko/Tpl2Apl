# (c) Николай Полягошко, 2015, 2016, 2017
# ver 1.2
#
#
# revhistory
#
# v.1.0
# Initial
#
# v.1.1
# пути источников и назначений теперь задаются в массивах
#
# v.1.1a
# имя лог файла включает дату
#
# v.1.2  11.02.2018
# Исправлена проверка успешности копирования


$Sources = @( "\\10.0.87.247\штрихкоды\*.tpl" )

$BackupPath = "\\10.0.87.247\штрихкоды\backup"
#$LogFile    = "\\labserver\LAB\Alisa\Tecan\tpl2apl.log"
$LogPath    = "\\labserver\LAB\Alisa\Tecan\Logs"
$ds = Get-Date -Format("yyyy_MM_dd")
$LogFile    = [IO.Path]::Combine($LogPath, "tpl2apl_" + $ds + ".log")

$Dests = @( "\\lab-tecan\c$\Users\Public\Documents\Tecan\Magellan\smp",
            "\\labtecan2\c$\Users\Public\Documents\Tecan\Magellan for F50\smp\",
            "\\labtecan3\c$\Users\Public\Documents\Tecan\Magellan for F50\smp\" )


if (Test-Path $LogFile) {
    "`n`n------------------------------------------`n" | Add-Content $LogFile   
} else {
    "" | Set-Content $LogFile
}
"$(Get-Date)`tНачинаем обработку TPL" | Add-Content $LogFile

Get-ChildItem $Sources | % {
    "`nОбрабатываем $_" | Add-Content $LogFile
    $an = $_.Name.Split('_.');

    [int]$plateId = $null;
    [int32]::TryParse($an[0], [ref]$plateId) | Out-Null;

    if ($an.Count -ge 4) {
        $plateId = [int]$an[1];
        $date = $an[2];
        $time = $an[3];
    }

    if (($plateId -eq $null) -or ($plateId -eq 0)) {[string]$plateId = "";}
    
    $filename = $null;

    Get-Content $_ | % {
        $a = $_.Split(';', [System.StringSplitOptions]::None);
        
        if ($a[0] -eq 'H') {
            if ($date -eq $null) {$date = a[1];}
            if ($time -eq $null) {$time = a[2];}
        }

        if ($a[0] -eq 'D') {

            [int]$sid = $null;
            if([int32]::TryParse($a[2], [ref]$sid) -eq $false) {return;}

            $method = $a[1].Replace(',', '');
            if ($method.Length -gt 12) {$method = $method.Substring(0, 12);}

            $well = $a[3];

            if ($date -ne $null) {$id2 = Get-Date $date -format "yyyyMMdd";}
            else {$id2 = Get-Date -format "yyyyMMdd";}

            $str = [string]::Format("{0,8}{1,4}   {2,12}{3,-12}{4,8}", $plateId, $well, $method, $sid, $id2);

            if ($filename -eq $null) {
                $filename = [string]::Format("{0} {1}_{2} {3}.apl",
                    $plateId, $date, $time.Replace('-','').Replace(':',''), $a[1].Replace('.','_'));

		$tmpfile = [System.IO.Path]::GetTempFileName();
                $str | Set-Content $tmpfile
            } else {
                $str | Add-Content $tmpfile
            }
        }
    }
    
    $Success = $False
    $Dests | % {
        $DestFile = Join-Path $_ $filename
        Copy-Item -Force $tmpfile $DestFile -ErrorAction SilentlyContinue
	"`n  Записываем $DestFile" | Add-Content $LogFile

        $DestExist = $False
        Try {
            $DestExist = Test-Path $DestFile -ErrorAction Stop
            if ($DestExist) { 
                $Success = $True
                "  Ок!" | Add-Content $LogFile
            }
            else { "  Неудача" | Add-Content $LogFile }
        }
        Catch { 
            "`n  Ошибка: $_.Exception.Message" | Add-Content $LogFile
        }
    }
    Remove-Item -Force $tmpfile

    if ($Success) {
    	Move-Item $_ "$BackupPath\$($_.Name).bak" -Force
    } else {
        "`nНи один из путей назначения недоступен. Обработка завершена с ошибкой" | Add-Content $LogFile
        exit 1
    }
}