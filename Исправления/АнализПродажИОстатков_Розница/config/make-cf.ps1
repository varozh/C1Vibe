# Собирает .cf Розницы с доработками на основе актуальной конфигурации рабочей базы:
#  - отчет БЛ_АнализПродажИОстатков + пункт в подменю «Отчеты для товароведов» (build-config.ps1);
#  - исправление печати QR в перемещении (..\..\ПечатьQR_ПеремещениеТоваров\apply-to-dump.ps1).
# Шаги: /DumpCfg и /DumpConfigToFiles рабочей базы -> /LoadCfg в новую файловую базу -> частичная загрузка
# из XML только измененных файлов (остальные объекты остаются побайтно как в рабочей базе, без пересборки
# через XML) -> проверка модулей -> /DumpCfg.
#   make-cf.ps1 -Work <рабочий каталог> -Cf <путь к .cf> [-Config mcp-retail\config.json]
param(
    [Parameter(Mandatory)][string]$Work,
    [Parameter(Mandatory)][string]$Cf,
    [string]$Config = 'D:\002_BASE\Vibe\mcp-retail\config.json',
    [string]$Platform = 'C:\Program Files (x86)\1cv8\8.3.12.1790\bin\1cv8.exe'
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$qr = Join-Path $here '..\..\ПечатьQR_ПеремещениеТоваров\apply-to-dump.ps1'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$sw = [Diagnostics.Stopwatch]::StartNew()
New-Item -ItemType Directory -Force $Work | Out-Null
function Run([string[]]$a, [string]$title) {
    $log = "$Work\$title.log"
    $p = Start-Process $Platform -ArgumentList ($a + @('/DisableStartupDialogs', '/DisableStartupMessages', '/Out', "`"$log`"")) -PassThru
    if (-not $p.WaitForExit(40 * 60 * 1000)) { try { $p.Kill() } catch {}; throw "$title`: конфигуратор не завершился за 40 мин" }
    $out = if (Test-Path $log) { (Get-Content $log -Encoding Default) -join "`n" } else { '' }
    Write-Host ("{0}: код {1}, {2:n0} c" -f $title, $p.ExitCode, $sw.Elapsed.TotalSeconds)
    return [pscustomobject]@{ Code = $p.ExitCode; Out = $out }
}

# 1. Рабочая база: .cf и XML
$cfg = Get-Content $Config -Raw -Encoding UTF8 | ConvertFrom-Json
$b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR((ConvertTo-SecureString ([string]$cfg.passwordEnc)))
try { $pwd1c = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
$srv = @('DESIGNER', '/S', "`"$($cfg.server)\$($cfg.ref)`"", '/N', "`"$($cfg.user)`"", '/P', "`"$pwd1c`"")
$liveCf = "$Work\live.cf"; $live = "$Work\live"
$r = Run ($srv + @('/DumpCfg', "`"$liveCf`"")) 'live-dumpcfg'; if ($r.Code -ne 0) { throw "выгрузка cf рабочей базы: $($r.Out)" }
if (Test-Path $live) { Remove-Item $live -Recurse -Force }
$r = Run ($srv + @('/DumpConfigToFiles', "`"$live`"")) 'live-dumpxml'; if ($r.Code -ne 0) { throw "выгрузка XML рабочей базы: $($r.Out)" }

# 2. XML доработок: отчет + интерфейс из свежей выгрузки, QR — правка модулей свежей выгрузки
& (Join-Path $here 'build-config.ps1') -DumpRoot $live
$part = "$Work\part"
if (Test-Path $part) { Remove-Item $part -Recurse -Force }
robocopy (Join-Path $here 'src') $part /E /NFL /NDL /NJH /NJS /NP | Out-Null
$qrFiles = 'CommonModules\УмКо_РаботаСМаркировкой.xml', 'CommonModules\УмКо_РаботаСМаркировкой\Ext\Module.bsl',
    'Documents\ПеремещениеТоваров\Forms\ФормаДокумента.xml', 'Documents\ПеремещениеТоваров\Forms\ФормаДокумента\Ext\Form.bin'
# Configuration.xml загружается вместе с корнем конфигурации: без файлов Ext (модули приложения, сеанса,
# внешнего соединения) частичная загрузка их удаляет. Подкаталоги Ext (Help, Splash) в список не берем —
# платформа их не принимает («Неизвестный объект метаданных Ext»), а справка и заставка при этом сохраняются.
# ParentConfigurations не трогаем.
$rootExt = Get-ChildItem "$live\Ext" -File | Where-Object { $_.Name -notmatch '^ParentConfigurations' } |
    ForEach-Object { $_.FullName.Substring($live.Length + 1) }
$qrFiles = @($qrFiles) + @($rootExt)
foreach ($f in $qrFiles) { New-Item -ItemType Directory -Force (Split-Path "$part\$f") | Out-Null; Copy-Item "$live\$f" "$part\$f" }
& $qr -DumpRoot $part
$files = Get-ChildItem $part -Recurse -File | ForEach-Object { $_.FullName.Substring($part.Length + 1) }
[IO.File]::WriteAllText("$Work\part.txt", ($files -join "`r`n"), $utf8Bom)
Write-Host "файлов для частичной загрузки: $($files.Count)"

# 3. Файловая база из .cf рабочей базы + частичная загрузка
$ib = "$Work\cfib"
if (Test-Path $ib) { Remove-Item $ib -Recurse -Force }
$p = Start-Process $Platform -ArgumentList @('CREATEINFOBASE', "File=`"$ib`";Locale=ru_RU", '/Out', "`"$Work\create.log`"") -Wait -PassThru
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/LoadCfg', "`"$liveCf`"") 'loadcfg'; if ($r.Code -ne 0) { throw "загрузка cf: $($r.Out)" }
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/LoadConfigFromFiles', "`"$part`"", '-listFile', "`"$Work\part.txt`"") 'partial'
Write-Host $r.Out
if ($r.Code -ne 0) { throw 'частичная загрузка не удалась' }
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/CheckModules', '-ThickClientOrdinaryApplication', '-Server', '-ExternalConnection') 'check'
$mine = $r.Out -split "`n" | Where-Object { $_ -match 'БЛ_АнализПродажИОстатков|УмКо_РаботаСМаркировкой|ПеремещениеТоваров' }
$other = ($r.Out -split "`n" | Where-Object { $_ -match '^\{' } | ForEach-Object { ($_ -split '\(')[0] } | Sort-Object -Unique) -join ', '
Write-Host "ошибок в доработанных модулях: $(if ($mine) { "`n" + ($mine -join "`n") } else { 'нет' }); модули с ошибками: $other"
if ($mine) { throw 'ошибки компиляции в доработке' }
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/DumpCfg', "`"$Cf`"") 'dumpcfg'; if ($r.Code -ne 0) { throw "выгрузка cf: $($r.Out)" }
$chk = "$Work\c"; if (Test-Path $chk) { Remove-Item $chk -Recurse -Force }
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/DumpConfigToFiles', "`"$chk`"") 'dumpcheck'
Write-Host "готово: $Cf"
