# Готовит доработку конфигурации УТ 10.3 (регламентное задание выгрузки остатков на BUSINESS LINE SMART PANEL)
# и проверяет её в тестовой файловой базе:
#   новый общий модуль БЛ_ПанельОстатковУТ (код: Доработка_УТ\БЛ_ПанельОстатковУТ.bsl)
#   + НЕпредопределённое регламентное задание БЛ_ВыгрузкаОстатковНаПанель.
# Экземпляр задания и расписание создаются в «Консоли заданий» (внешняя обработка УТ0000044) — как в Рознице,
# где задание БЛ_ВыгрузкаНаПанель уже есть в конфигурации и доработки не требует.
# Расширения в УТ невозможны: режим совместимости конфигурации 8.2.16.
#
# Результат: Доработка_УТ\src — XML для частичной загрузки в тестовую базу (список файлов — files.txt),
#            Доработка_УТ\*.cf  — конфигурация ТЕСТОВОЙ базы с доработкой (только для проверки; в рабочую базу не загружать — выгрузка устарела).
#
#   build-config.ps1 -UtBase <файловая база с конфигурацией УТ>
# Тестовая база: 1cv8 CREATEINFOBASE File=<путь>; 1cv8 DESIGNER /F <путь> /LoadConfigFromFiles D:\002_BASE\Vibe\УТ /UpdateDBCfg
param(
    [string]$UtBase,
    [string]$Platform = 'C:\Program Files (x86)\1cv8\8.3.12.1790\bin\1cv8.exe',
    [string]$DumpRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$utf8Bom = New-Object System.Text.UTF8Encoding($true)

function Get-Uuid([string]$s) {
    $h = [Security.Cryptography.MD5]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes("BL-panel-cfg:" + $s))
    return ([guid]::new($h)).ToString()
}
function Save([string]$path, [string]$text) {
    New-Item -ItemType Directory -Force (Split-Path -Parent $path) | Out-Null
    [IO.File]::WriteAllText($path, $text.Replace("`r`n", "`n").Replace("`n", "`r`n"), $utf8Bom)
}
function Read-Text([string]$path) { [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) }

$ns = 'xmlns="http://v8.1c.ru/8.3/MDClasses" xmlns:app="http://v8.1c.ru/8.2/managed-application/core" xmlns:cfg="http://v8.1c.ru/8.1/data/enterprise/current-config" xmlns:cmi="http://v8.1c.ru/8.2/managed-application/cmi" xmlns:ent="http://v8.1c.ru/8.1/data/enterprise" xmlns:lf="http://v8.1c.ru/8.2/managed-application/logform" xmlns:style="http://v8.1c.ru/8.1/data/ui/style" xmlns:sys="http://v8.1c.ru/8.1/data/ui/fonts/system" xmlns:v8="http://v8.1c.ru/8.1/data/core" xmlns:v8ui="http://v8.1c.ru/8.1/data/ui" xmlns:web="http://v8.1c.ru/8.1/data/ui/colors/web" xmlns:win="http://v8.1c.ru/8.1/data/ui/colors/windows" xmlns:xen="http://v8.1c.ru/8.3/xcf/enums" xmlns:xpr="http://v8.1c.ru/8.3/xcf/predef" xmlns:xr="http://v8.1c.ru/8.3/xcf/readable" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="2.6"'

function Job-Xml($uuid, $name, $synonym, $comment, $method, [bool]$predefined = $false) {
    $pre = $predefined.ToString().ToLower()
@"
<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject $ns>
	<ScheduledJob uuid="$uuid">
		<Properties>
			<Name>$name</Name>
			<Synonym>
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>$synonym</v8:content>
				</v8:item>
			</Synonym>
			<Comment>$comment</Comment>
			<MethodName>$method</MethodName>
			<Description>BL Smart Panel</Description>
			<Key/>
			<Use>true</Use>
			<Predefined>$pre</Predefined>
			<RestartCountOnFailure>3</RestartCountOnFailure>
			<RestartIntervalOnFailure>600</RestartIntervalOnFailure>
		</Properties>
	</ScheduledJob>
</MetaDataObject>
"@
}

# Добавляет строку <Kind>Name</Kind> в ChildObjects корневого Configuration.xml после последнего элемента того же вида.
function Add-Child([string]$xml, [string]$kind, [string]$name) {
    if ($xml -match "<$kind>$([regex]::Escape($name))</$kind>") { return $xml }
    $m = [regex]::Matches($xml, "(?m)^(\s*)<$kind>[^<]*</$kind>\r?\n")
    if ($m.Count -eq 0) { throw "В Configuration.xml нет элементов вида $kind" }
    $last = $m[$m.Count - 1]
    $line = "$($last.Groups[1].Value)<$kind>$name</$kind>`r`n"
    return $xml.Insert($last.Index + $last.Length, $line)
}

function Invoke-Designer($base, $argLine, $title) {
    $log = Join-Path $env:TEMP 'bl-build-config.log'
    Remove-Item $log -ErrorAction SilentlyContinue
    $p = Start-Process $Platform -ArgumentList "DESIGNER /F `"$base`" /DisableStartupDialogs /DisableStartupMessages $argLine /Out `"$log`"" -PassThru
    if (-not $p.WaitForExit(30 * 60 * 1000)) { try { $p.Kill() } catch {}; throw "$title`: конфигуратор не завершился за 30 мин." }
    $out = if (Test-Path $log) { (Get-Content $log -Encoding Default) -join "`n" } else { '' }
    return [pscustomobject]@{ Code = $p.ExitCode; Out = $out }
}

function Build($d) {
    $dir = Join-Path $here $d.Folder
    $src = Join-Path $dir 'src'
    if (Test-Path $src) { Remove-Item $src -Recurse -Force }
    & $d.Generate $src
    $files = Get-ChildItem $src -Recurse -File | ForEach-Object { $_.FullName.Substring($src.Length + 1) }
    Save (Join-Path $dir 'files.txt') ($files -join "`n")
    Write-Host "[$($d.Name)] XML для частичной загрузки: $src ($($files.Count) файлов)"
    if (-not $d.Base) { return }

    $r = Invoke-Designer $d.Base "/LoadConfigFromFiles `"$src`" -listFile `"$(Join-Path $dir 'files.txt')`"" 'Загрузка'
    Write-Host "[$($d.Name)] загрузка -> $($r.Code) $($r.Out)"
    if ($r.Code -ne 0) { throw "$($d.Name): частичная загрузка не удалась" }
    $r = Invoke-Designer $d.Base '/UpdateDBCfg' 'Обновление'
    Write-Host "[$($d.Name)] обновление БД -> $($r.Code) $($r.Out)"
    if ($r.Code -ne 0) { throw "$($d.Name): обновление конфигурации БД не удалось" }
    $r = Invoke-Designer $d.Base '/CheckModules -Server -ExternalConnection' 'Проверка модулей'
    $mine = ($r.Out -split "`n" | Where-Object { $_ -match $d.CheckPattern }) -join "`n"
    Write-Host "[$($d.Name)] проверка модулей -> $($r.Code); ошибок в доработке: $(if ($mine) { "`n$mine" } else { 'нет' })"
    if ($mine) { throw "$($d.Name): ошибки компиляции в доработке" }
    $cf = Join-Path $dir "$($d.Name).cf"
    $r = Invoke-Designer $d.Base "/DumpCfg `"$cf`"" 'Выгрузка cf'
    if ($r.Code -ne 0) { throw "$($d.Name): не удалось выгрузить cf: $($r.Out)" }
    Write-Host "[$($d.Name)] готово: $cf"
}

$ut = [pscustomobject]@{
    Name = 'УТ_БЛ_ВыгрузкаОстатковНаПанель'; Folder = 'Доработка_УТ'; Base = $UtBase
    CheckPattern = 'БЛ_ПанельОстатковУТ|БЛ_ВыгрузкаОстатковНаПанель'
    Generate = {
        param($src)
        $cfgXml = Read-Text (Join-Path $DumpRoot 'УТ\Configuration.xml')
        $cfgXml = Add-Child $cfgXml 'CommonModule' 'БЛ_ПанельОстатковУТ'
        $cfgXml = Add-Child $cfgXml 'ScheduledJob' 'БЛ_ВыгрузкаОстатковНаПанель'
        Save "$src\Configuration.xml" $cfgXml

        Save "$src\CommonModules\БЛ_ПанельОстатковУТ.xml" @"
<?xml version="1.0" encoding="UTF-8"?>
<MetaDataObject $ns>
	<CommonModule uuid="$(Get-Uuid 'ut:module')">
		<Properties>
			<Name>БЛ_ПанельОстатковУТ</Name>
			<Synonym>
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>BL: остатки складов на панель</v8:content>
				</v8:item>
			</Synonym>
			<Comment>Выгрузка остатков складов на BUSINESS LINE SMART PANEL (регламентное задание БЛ_ВыгрузкаОстатковНаПанель)</Comment>
			<Global>false</Global>
			<ClientManagedApplication>false</ClientManagedApplication>
			<Server>true</Server>
			<ExternalConnection>true</ExternalConnection>
			<ClientOrdinaryApplication>false</ClientOrdinaryApplication>
			<ServerCall>false</ServerCall>
			<Privileged>false</Privileged>
			<ReturnValuesReuse>DontUse</ReturnValuesReuse>
		</Properties>
	</CommonModule>
</MetaDataObject>
"@
        New-Item -ItemType Directory -Force "$src\CommonModules\БЛ_ПанельОстатковУТ\Ext" | Out-Null
        Copy-Item (Join-Path $here 'Доработка_УТ\БЛ_ПанельОстатковУТ.bsl') "$src\CommonModules\БЛ_ПанельОстатковУТ\Ext\Module.bsl"

        Save "$src\ScheduledJobs\БЛ_ВыгрузкаОстатковНаПанель.xml" (Job-Xml (Get-Uuid 'ut:job') 'БЛ_ВыгрузкаОстатковНаПанель' `
            'BL: остатки складов на панель' 'Остатки складов УТ на SMART PANEL 10.28.2.164:8086. Экземпляр и расписание (рекомендуется 08:00, 13:00, 18:00, 21:00) - в Консоли заданий.' `
            'CommonModule.БЛ_ПанельОстатковУТ.ВыгрузкаОстатковНаПанель' $false)
    }
}

Build $ut
