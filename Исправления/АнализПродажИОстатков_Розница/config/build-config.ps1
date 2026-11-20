# Готовит доработку конфигурации Розница 1.0: отчет БЛ_АнализПродажИОстатков — копия внешнего отчета
# ..\src\АнализПродажИОстатков (СКД + модуль объекта), синоним с «(+)», пункт последней строкой подменю
# «Отчеты для товароведов» интерфейса «Полный». Явных прав в ролях не добавляет: отчеты этого подменю
# доступны ролям с правами на новые объекты (ПолныеПрава и т. п.), как и остальные отчеты подменю.
#
#   build-config.ps1 [-Work <рабочий каталог>]
# Без -Work только генерирует config\src и files.txt. С -Work: полная загрузка копии выгрузки с файлами
# config\src поверх в новую файловую базу (интерфейс на поддержке) и проверка модулей.
param(
    [string]$Work,
    [string]$Platform = 'C:\Program Files (x86)\1cv8\8.3.12.1790\bin\1cv8.exe',
    [string]$DumpRoot = 'D:\002_BASE\Vibe\Розница'
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$ExtName = 'АнализПродажИОстатков'
$ext = Join-Path (Split-Path -Parent $here) "src\$ExtName"
$src = Join-Path $here 'src'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)

$Name = 'БЛ_АнализПродажИОстатков'
$Synonym = 'Анализ продаж и остатков (цены на 5 и на 0) (+)'
$Interface = 'Полный'
$AfterReport = 'Отчет.ПродажаТоваровВКомплекте'   # новый пункт — сразу после него (последняя строка подменю)

function Get-Uuid([string]$s) {
    $h = [Security.Cryptography.MD5]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes("BL-rt-sales-stock:" + $s))
    return ([guid]::new($h)).ToString()
}
function Save([string]$path, [string]$text) {
    New-Item -ItemType Directory -Force (Split-Path -Parent $path) | Out-Null
    [IO.File]::WriteAllText($path, $text, $utf8Bom)
}
function Read-Text([string]$path) { [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) }

if (Test-Path $src) { Remove-Item $src -Recurse -Force }

$U = @{
    Report = Get-Uuid 'report'; ObjType = Get-Uuid 'obj-type'; ObjValue = Get-Uuid 'obj-value'
    MgrType = Get-Uuid 'mgr-type'; MgrValue = Get-Uuid 'mgr-value'; MenuItem = Get-Uuid 'menu-item'
    Template = Get-Uuid 'template'
}

# --- 1. Отчет
$extXml = Read-Text "$ext.xml"
$hdr = $extXml.Substring(0, $extXml.IndexOf('<ExternalReport'))
$tplName = [regex]::Match($extXml, '<Template>([^<]+)</Template>').Groups[1].Value
Save "$src\Reports\$Name.xml" ($hdr + @"
<Report uuid="$($U.Report)">
		<InternalInfo>
			<xr:GeneratedType name="ReportObject.$Name" category="Object">
				<xr:TypeId>$($U.ObjType)</xr:TypeId>
				<xr:ValueId>$($U.ObjValue)</xr:ValueId>
			</xr:GeneratedType>
			<xr:GeneratedType name="ReportManager.$Name" category="Manager">
				<xr:TypeId>$($U.MgrType)</xr:TypeId>
				<xr:ValueId>$($U.MgrValue)</xr:ValueId>
			</xr:GeneratedType>
		</InternalInfo>
		<Properties>
			<Name>$Name</Name>
			<Synonym>
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>$Synonym</v8:content>
				</v8:item>
			</Synonym>
			<Comment>BL: копия внешнего отчета «Анализ продаж и остатков (цены на 5 и на 0)»</Comment>
			<UseStandardCommands>false</UseStandardCommands>
			<DefaultForm/>
			<AuxiliaryForm/>
			<MainDataCompositionSchema>Report.$Name.Template.$tplName</MainDataCompositionSchema>
			<DefaultSettingsForm/>
			<AuxiliarySettingsForm/>
			<DefaultVariantForm/>
			<VariantsStorage/>
			<SettingsStorage/>
			<IncludeHelpInContents>false</IncludeHelpInContents>
			<ExtendedPresentation/>
			<Explanation/>
		</Properties>
		<ChildObjects>
			<Template>$tplName</Template>
		</ChildObjects>
	</Report>
</MetaDataObject>
"@)

$tx = [regex]::Replace((Read-Text "$ext\Templates\$tplName.xml"), 'Template uuid="[^"]+"', "Template uuid=`"$($U.Template)`"")
Save "$src\Reports\$Name\Templates\$tplName.xml" $tx
New-Item -ItemType Directory -Force "$src\Reports\$Name\Templates\$tplName\Ext" | Out-Null
Copy-Item "$ext\Templates\$tplName\Ext\Template.xml" "$src\Reports\$Name\Templates\$tplName\Ext\Template.xml"
$module = Read-Text "$ext\Ext\ObjectModule.bsl"
Save "$src\Reports\$Name\Ext\ObjectModule.bsl" $module.TrimStart([char]0xFEFF)
Save (Join-Path $here "$($Name)_МодульОбъекта.bsl") $module.TrimStart([char]0xFEFF)

# --- 2. Configuration.xml: новый отчет после последнего отчета
$cfg = Read-Text "$DumpRoot\Configuration.xml"
if ($cfg -notmatch "<Report>$Name</Report>") {
    $m = [regex]::Matches($cfg, "(?m)^(\s*)<Report>[^<]*</Report>\r?\n")
    $last = $m[$m.Count - 1]
    $cfg = $cfg.Insert($last.Index + $last.Length, "$($last.Groups[1].Value)<Report>$Name</Report>`r`n")
}
Save "$src\Configuration.xml" $cfg.TrimStart([char]0xFEFF)

# --- 3. Интерфейс «Полный»: команда «Открыть отчет» + пункт подменю после «Продажи товаров в комплекте»
$it = (Read-Text "$DumpRoot\Interfaces\$Interface\Ext\Interface.bin").TrimStart([char]0xFEFF)
$h = [regex]::Match($it, '^\{3,\r?\n\{5,([0-9a-f-]{36}),(\d+),0,(\d+),')
if (-not $h.Success) { throw 'Interface.bin: не разобран заголовок' }
$cls = $h.Groups[1].Value; $maxId = [int]$h.Groups[2].Value; $cmdCount = [int]$h.Groups[3].Value
$newId = $maxId + 1
$it = $it.Substring(0, $h.Index) + $h.Value.Replace(",$maxId,0,$cmdCount,", ",$newId,0,$($cmdCount + 1),") + $it.Substring($h.Index + $h.Length)
$cm = [regex]::Match($it, '\{6,([0-9a-f-]{36}),1,([0-9a-f-]{36}),\r?\n\{6,([0-9a-f-]{36}),([0-9a-f-]{36}),0,0,1,\r?\n\{"U"\},\r?\n\{1,0,[^\r\n]*\r?\n\{1,"' + [regex]::Escape($AfterReport) + ': Открыть отчет",[\s\S]*?\r?\n\{0,0,0\}\r?\n\}\r?\n\},(\d+),\r?\n\{1,1,\r?\n\{"ru","[^"]*"\}\r?\n\}\r?\n\},\r?\n')
if (-not $cm.Success) { throw "Interface.bin: не найдена команда $AfterReport" }
$oldItem = $cm.Groups[1].Value; $oldObj = $cm.Groups[3].Value
$oldSyn = [regex]::Match($cm.Value, '\{"ru","([^"]*)"\}').Groups[1].Value
$newCmd = $cm.Value.Replace($oldItem, $U.MenuItem).Replace($oldObj, $U.Report).Replace("$AfterReport`: Открыть отчет", "Отчет.$Name`: Открыть отчет").Replace("`"$oldSyn`"", "`"$Synonym`"").Replace("Открыть отчет $oldSyn", "Открыть отчет $Synonym")
$it = $it.Insert($cm.Index + $cm.Length, $newCmd)
$pname = 'Отчет_' + $AfterReport.Split('.')[1] + '_ОткрытьОтчет'
$mi = [regex]::Match($it, [regex]::Escape($oldItem) + ',\r?\n\{8,"' + $pname + '",0,(\d+),\r?\n\{1,1,\r?\n\{"ru","[^"]*"\}\r?\n\},(\d),' + $cls + ',(\d+),(1e2,0,0,1,0,\d,0,0\})')
if (-not $mi.Success) { throw 'Interface.bin: не найден пункт меню' }
$nl = if ($mi.Value.Contains("`r`n")) { "`r`n" } else { "`n" }
$newMi = ",$($U.MenuItem),$nl{8,`"Отчет_$($Name)_ОткрытьОтчет`",0,$($mi.Groups[1].Value),$nl{1,1,$nl{`"ru`",`"$Synonym`"}$nl},$($mi.Groups[2].Value),$cls,$newId,$($mi.Groups[4].Value)"
$it = $it.Insert($mi.Index + $mi.Length, $newMi)
$menuHead = [regex]::Matches($it.Substring(0, $mi.Index), '\{5,([0-9a-f-]{36}),(\d+),0,(\d+),') | Select-Object -Last 1
$cnt = [int]$menuHead.Groups[3].Value
$it = $it.Substring(0, $menuHead.Index) + $menuHead.Value.Replace(",0,$cnt,", ",0,$($cnt + 1),") + $it.Substring($menuHead.Index + $menuHead.Length)
Save "$src\Interfaces\$Interface\Ext\Interface.bin" $it
Copy-Item "$DumpRoot\Interfaces\$Interface.xml" "$src\Interfaces\$Interface.xml"
Write-Host "Интерфейс: команд $cmdCount -> $($cmdCount + 1), id пункта $newId, подменю $($menuHead.Groups[2].Value): пунктов $cnt -> $($cnt + 1)"

$files = Get-ChildItem $src -Recurse -File | ForEach-Object { $_.FullName.Substring($src.Length + 1) }
[IO.File]::WriteAllText((Join-Path $here 'files.txt'), ($files -join "`r`n"), $utf8Bom)
Write-Host "XML доработки: $src ($($files.Count) файлов)"
if (-not $Work) { return }

# --- 4. Проверка: полная загрузка копии выгрузки + доработка в новую файловую базу
$sw = [Diagnostics.Stopwatch]::StartNew()
function Run([string[]]$a, [string]$title) {
    $log = "$Work\$title.log"
    $p = Start-Process $Platform -ArgumentList ($a + @('/Out', "`"$log`"")) -Wait -PassThru
    $out = if (Test-Path $log) { (Get-Content $log -Encoding Default) -join "`n" } else { '' }
    Write-Host ("{0}: код {1}, {2:n0} c" -f $title, $p.ExitCode, $sw.Elapsed.TotalSeconds)
    return [pscustomobject]@{ Code = $p.ExitCode; Out = $out }
}
$dump = "$Work\rtdump"; $ib = "$Work\rtcfib"
robocopy $DumpRoot $dump /MIR /NFL /NDL /NJH /NJS /NP /R:1 /W:1 | Out-Null
robocopy $src $dump /E /NFL /NDL /NJH /NJS /NP | Out-Null
if (Test-Path $ib) { Remove-Item $ib -Recurse -Force }
$r = Run @('CREATEINFOBASE', "File=`"$ib`";Locale=ru_RU") 'create'
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/DisableStartupDialogs', '/DisableStartupMessages', '/LoadConfigFromFiles', "`"$dump`"") 'load'
$r.Out -split "`n" | Where-Object { $_ -notmatch 'неверная ссылка' } | Select-Object -Last 5
if ($r.Code -ne 0) { throw 'загрузка не удалась' }
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/DisableStartupDialogs', '/DisableStartupMessages', '/CheckModules', '-ThickClientOrdinaryApplication', '-Server', '-ExternalConnection') 'check'
$mine = $r.Out -split "`n" | Where-Object { $_ -match $Name }
Write-Host "ошибок в $Name`: $(if ($mine) { "`n" + ($mine -join "`n") } else { 'нет' })"
if ($mine) { throw 'ошибки компиляции в отчете' }
$r = Run @('DESIGNER', '/F', "`"$ib`"", '/DisableStartupDialogs', '/DisableStartupMessages', '/DumpConfigToFiles', "`"$Work\rtcheck`"") 'dumpcheck'
Write-Host 'проверка в тестовой базе завершена'
