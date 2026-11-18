# Собирает УТ_ОтчетПоКладовщикам.cf: полная загрузка копии выгрузки УТ с файлами config\src поверх
# в новую файловую базу (частичная загрузка невозможна — отчет на поддержке), проверка модулей, /DumpCfg.
#   make-cf.ps1 -Work <рабочий каталог для копии выгрузки и базы>
param([Parameter(Mandatory)][string]$Work)
$ErrorActionPreference = 'Stop'
$exe = 'C:\Program Files (x86)\1cv8\8.3.12.1790\bin\1cv8.exe'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$dump = "$Work\utdump"; $ib = "$Work\utcfib"
$sw = [Diagnostics.Stopwatch]::StartNew()
function Run([string[]]$a, [string]$title) {
    $log = "$Work\$title.log"
    $p = Start-Process $exe -ArgumentList ($a + @('/Out', "`"$log`"")) -Wait -PassThru
    $out = if (Test-Path $log) { (Get-Content $log -Encoding Default) -join "`n" } else { '' }
    "{0}: код {1}, {2:n0} c" -f $title, $p.ExitCode, $sw.Elapsed.TotalSeconds
    return [pscustomobject]@{ Code = $p.ExitCode; Out = $out }
}
robocopy 'D:\002_BASE\Vibe\УТ' $dump /MIR /NFL /NDL /NJH /NJS /NP /R:1 /W:1 | Out-Null
robocopy "$here\src" $dump /E /NFL /NDL /NJH /NJS /NP | Out-Null
"копия выгрузки {0:n0} c" -f $sw.Elapsed.TotalSeconds
if (Test-Path $ib) { Remove-Item $ib -Recurse -Force }
$r = Run @('CREATEINFOBASE', "File=`"$ib`";Locale=ru_RU") 'create'
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/DisableStartupDialogs', '/DisableStartupMessages', '/LoadConfigFromFiles', "`"$dump`"") 'load'
$r.Out -split "`n" | Where-Object { $_ -notmatch 'неверная ссылка' } | Select-Object -Last 5
if ($r.Code -ne 0) { throw 'загрузка не удалась' }
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/DisableStartupDialogs', '/DisableStartupMessages', '/CheckModules', '-ThickClientOrdinaryApplication', '-Server', '-ExternalConnection') 'check'
$mine = $r.Out -split "`n" | Where-Object { $_ -match 'ОтчетПоКладовщикам' }
"ошибок в ОтчетПоКладовщикам: $(if ($mine) { "`n" + ($mine -join "`n") } else { 'нет' })"
if ($mine) { throw 'ошибки компиляции в отчете' }
$cf = Join-Path (Split-Path -Parent $here) 'УТ_ОтчетПоКладовщикам.cf'
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/DisableStartupDialogs', '/DisableStartupMessages', '/DumpCfg', "`"$cf`"") 'dumpcfg'
if ($r.Code -ne 0) { throw "выгрузка cf не удалась: $($r.Out)" }
"готово: $cf"
