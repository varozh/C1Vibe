# Тестовая база УТ без поддержки: копия выгрузки без Ext\ParentConfigurations.bin + файлы доработки поверх.
param([Parameter(Mandatory)][string]$S)   # рабочий каталог: туда кладутся копия выгрузки и база utib2
$ErrorActionPreference = 'Stop'
$exe = 'C:\Program Files (x86)\1cv8\8.3.12.1790\bin\1cv8.exe'
$dump = "$S\utdump"
$add = 'D:\002_BASE\Vibe\Исправления\АнализОстатковСоШтрихкодами_УТ\config\src'
$sw = [Diagnostics.Stopwatch]::StartNew()
robocopy 'D:\002_BASE\Vibe\УТ' $dump /MIR /NFL /NDL /NJH /NJS /NP /R:1 /W:1 | Out-Null
"copy done {0:n0}s" -f $sw.Elapsed.TotalSeconds
# полная загрузка не проверяет запреты поддержки, поэтому роли и интерфейс загружаются и при наличии Ext\ParentConfigurations
Get-ChildItem "$dump\Ext" | Select-Object Name
robocopy $add $dump /E /NFL /NDL /NJH /NJS /NP | Out-Null
$ib = "$S\utib2"
if (Test-Path $ib) { Remove-Item $ib -Recurse -Force }
$p = Start-Process $exe -ArgumentList @('CREATEINFOBASE', "File=`"$ib`";Locale=ru_RU", '/Out', "`"$S\create2.log`"") -Wait -PassThru
"create $($p.ExitCode)"
$p = Start-Process $exe -ArgumentList @('DESIGNER','/F',"`"$ib`"",'/DisableStartupDialogs','/DisableStartupMessages','/LoadConfigFromFiles',"`"$dump`"",'/Out',"`"$S\load2.log`"") -Wait -PassThru
"load $($p.ExitCode) {0:n0}s" -f $sw.Elapsed.TotalSeconds
Get-Content "$S\load2.log" -Encoding Default | Where-Object { $_ -notmatch 'неверная ссылка' } | Select-Object -Last 10
