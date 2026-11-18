# Готовит доработку конфигурации УТ 10.3: отчет БЛ_АнализОстатковТоваровНаСкладахШК
# (копия внешнего отчета ..\src, остатки + ШК 46 организаций) и пункт «Анализ остатков товаров на складах (ШК 46)»
# последней строкой меню «Отчеты» интерфейса «Полный»; права Use/View — тем же ролям, что у
# «Анализа доступности товаров на складах» (у полных ролей права на новые объекты ставятся автоматически).
#
#   build-config.ps1 [-UtBase <файловая база с конфигурацией УТ>]
# Без -UtBase только генерирует src и files.txt. С -UtBase: частичная загрузка и проверка модулей в тестовой базе.
param(
    [string]$UtBase,
    [string]$Platform = 'C:\Program Files (x86)\1cv8\8.3.12.1790\bin\1cv8.exe',
    [string]$DumpRoot = 'D:\002_BASE\Vibe\УТ'
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$ext = Join-Path (Split-Path -Parent $here) 'src\АнализОстатковТоваровНаСкладахШК'
$src = Join-Path $here 'src'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)

$Name = 'БЛ_АнализОстатковТоваровНаСкладахШК'
$Synonym = 'Анализ остатков товаров на складах (ШК 46)'
$Roles = 'ЗаведующийСкладом', 'Кладовщик', 'МенеджерПоЗакупкам', 'МенеджерПоПродажам'
$Interface = 'Полный'
$AfterReport = 'Отчет.ОтчетПоКладовщикам'   # новый пункт — сразу после него (последняя строка меню «Отчеты»)

# Детерминированные uuid: повторная сборка дает те же идентификаторы.
function Get-Uuid([string]$s) {
    $h = [Security.Cryptography.MD5]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes("BL-ut-stock-barcodes:" + $s))
    return ([guid]::new($h)).ToString()
}
function Save([string]$path, [string]$text) {
    New-Item -ItemType Directory -Force (Split-Path -Parent $path) | Out-Null
    [IO.File]::WriteAllText($path, $text, $utf8Bom)
}
function Read-Text([string]$path) { [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) }
function Replace-Once([string]$text, [string]$old, [string]$new, [string]$what) {
    $i = $text.IndexOf($old, [StringComparison]::Ordinal)
    if ($i -lt 0) { throw "Не найдено: $what" }
    if ($text.IndexOf($old, $i + 1, [StringComparison]::Ordinal) -ge 0) { throw "Найдено больше одного: $what" }
    return $text.Substring(0, $i) + $new + $text.Substring($i + $old.Length)
}
# Замена ASCII-строки одинаковой длины в бинарном файле (Form.bin — контейнер, длины блоков не меняются).
function Replace-Bytes([byte[]]$b, [string]$old, [string]$new) {
    if ($old.Length -ne $new.Length) { throw "Разная длина: $old / $new" }
    $o = [Text.Encoding]::ASCII.GetBytes($old); $n = [Text.Encoding]::ASCII.GetBytes($new); $cnt = 0
    for ($i = 0; $i -le $b.Length - $o.Length; $i++) {
        $ok = $true
        for ($k = 0; $k -lt $o.Length; $k++) { if ($b[$i + $k] -ne $o[$k]) { $ok = $false; break } }
        if ($ok) { [Array]::Copy($n, 0, $b, $i, $n.Length); $cnt++ }
    }
    return $cnt
}

if (Test-Path $src) { Remove-Item $src -Recurse -Force }

# --- uuid нового отчета
$U = @{
    Report = Get-Uuid 'report'; ObjType = Get-Uuid 'obj-type'; ObjValue = Get-Uuid 'obj-value'
    MgrType = Get-Uuid 'mgr-type'; MgrValue = Get-Uuid 'mgr-value'; Attr = Get-Uuid 'attr'; Form = Get-Uuid 'form'
    MenuItem = Get-Uuid 'menu-item'
}

# --- 1. Отчет: метаданные из внешнего отчета (форма, реквизит УниверсальныйОтчет, модуль)
$extXml = Read-Text "$ext.xml"
$hdr = $extXml.Substring(0, $extXml.IndexOf('<ExternalReport'))
$attr = $extXml.Substring($extXml.IndexOf('<Attribute uuid'), $extXml.IndexOf('</Attribute>') + 12 - $extXml.IndexOf('<Attribute uuid'))
$extAttrUuid = [regex]::Match($attr, 'Attribute uuid="([^"]+)"').Groups[1].Value
$extTypeId = [regex]::Match($extXml, '<xr:TypeId>([^<]+)</xr:TypeId>').Groups[1].Value
$attr = $attr.Replace("Attribute uuid=`"$extAttrUuid`"", "Attribute uuid=`"$($U.Attr)`"")

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
			<Comment>BL: остатки на складах + штрихкоды 46 организаций (копия «Анализа доступности товаров на складах»)</Comment>
			<UseStandardCommands>false</UseStandardCommands>
			<DefaultForm>Report.$Name.Form.Форма</DefaultForm>
			<AuxiliaryForm/>
			<MainDataCompositionSchema/>
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
			$attr
			<Form>Форма</Form>
		</ChildObjects>
	</Report>
</MetaDataObject>
"@)

$formXml = (Read-Text "$ext\Forms\Форма.xml") -replace 'Form uuid="[^"]+"', "Form uuid=`"$($U.Form)`""
Save "$src\Reports\$Name\Forms\Форма.xml" $formXml

# Form.bin: тип основного реквизита «ОтчетОбъект» -> новый отчет; ссылки на реквизит УниверсальныйОтчет -> новый uuid.
$bin = [IO.File]::ReadAllBytes("$ext\Forms\Форма\Ext\Form.bin")
$n1 = Replace-Bytes $bin $extTypeId $U.ObjType
$n2 = Replace-Bytes $bin $extAttrUuid $U.Attr
if ($n1 -ne 1 -or $n2 -lt 1) { throw "Form.bin: замен типа $n1, реквизита $n2" }
New-Item -ItemType Directory -Force "$src\Reports\$Name\Forms\Форма\Ext" | Out-Null
[IO.File]::WriteAllBytes("$src\Reports\$Name\Forms\Форма\Ext\Form.bin", $bin)

$module = Read-Text "$ext\Ext\ObjectModule.bsl"
$module = $module.Replace('// Копия отчета конфигурации УТ 10.3 «Анализ доступности товаров на складах».',
    "// Отчет $Name — копия отчета конфигурации УТ 10.3 «Анализ доступности товаров на складах».")
Save "$src\Reports\$Name\Ext\ObjectModule.bsl" $module

# --- 2. Configuration.xml: новый отчет после последнего отчета
$cfg = Read-Text "$DumpRoot\Configuration.xml"
if ($cfg -notmatch "<Report>$Name</Report>") {
    $m = [regex]::Matches($cfg, "(?m)^(\s*)<Report>[^<]*</Report>\r?\n")
    $last = $m[$m.Count - 1]
    $cfg = $cfg.Insert($last.Index + $last.Length, "$($last.Groups[1].Value)<Report>$Name</Report>`r`n")
}
Save "$src\Configuration.xml" $cfg

# --- 3. Права: Use + View в ролях
foreach ($r in $Roles) {
    $x = Read-Text "$DumpRoot\Roles\$r\Ext\Rights.xml"
    if ($x -notmatch "<name>Report\.$Name</name>") {
        $obj = "`t<object>`r`n`t`t<name>Report.$Name</name>`r`n`t`t<right>`r`n`t`t`t<name>Use</name>`r`n`t`t`t<value>true</value>`r`n`t`t</right>`r`n`t`t<right>`r`n`t`t`t<name>View</name>`r`n`t`t`t<value>true</value>`r`n`t`t</right>`r`n`t</object>`r`n"
        $i = $x.IndexOf("`t<object>", [StringComparison]::Ordinal)
        $x = $x.Insert($i, $obj)
    }
    Save "$src\Roles\$r\Ext\Rights.xml" $x
    Copy-Item "$DumpRoot\Roles\$r.xml" "$src\Roles\$r.xml"
}

# --- 4. Интерфейс «Полный»: команда «Открыть отчет» + пункт меню «Отчеты» после «Отчета по кладовщикам»
$it = Read-Text "$DumpRoot\Interfaces\$Interface\Ext\Interface.bin"
# 4.1 заголовок: {5,<id>,<счетчик идентификаторов>,0,<число команд>,
$h = [regex]::Match($it, '^\uFEFF?\{3,\r?\n\{5,([0-9a-f-]{36}),(\d+),0,(\d+),')
if (-not $h.Success) { throw 'Interface.bin: не разобран заголовок' }
$cls = $h.Groups[1].Value; $maxId = [int]$h.Groups[2].Value; $cmdCount = [int]$h.Groups[3].Value
$newId = $maxId + 1
$it = $it.Substring(0, $h.Index) + $h.Value.Replace(",$maxId,0,$cmdCount,", ",$newId,0,$($cmdCount + 1),") + $it.Substring($h.Index + $h.Length)
# 4.2 команда: копия команды «Отчет по кладовщикам»
$cm = [regex]::Match($it, '\{6,([0-9a-f-]{36}),1,([0-9a-f-]{36}),\r?\n\{6,([0-9a-f-]{36}),([0-9a-f-]{36}),0,0,1,\r?\n\{"U"\},\r?\n\{1,0,[^\r\n]*\r?\n\{1,"' + [regex]::Escape($AfterReport) + ': Открыть отчет",[\s\S]*?\r?\n\{0,0,0\}\r?\n\}\r?\n\},(\d+),\r?\n\{1,1,\r?\n\{"ru","[^"]*"\}\r?\n\}\r?\n\},\r?\n')
if (-not $cm.Success) { throw "Interface.bin: не найдена команда $AfterReport" }
$oldItem = $cm.Groups[1].Value; $oldObj = $cm.Groups[3].Value
$oldSyn = [regex]::Match($cm.Value, '\{"ru","([^"]*)"\}').Groups[1].Value
$newCmd = $cm.Value.Replace($oldItem, $U.MenuItem).Replace($oldObj, $U.Report).Replace("$AfterReport`: Открыть отчет", "Отчет.$Name`: Открыть отчет").Replace("`"$oldSyn`"", "`"$Synonym`"").Replace("Открыть отчет $oldSyn", "Открыть отчет $Synonym")
$it = $it.Insert($cm.Index + $cm.Length, $newCmd)
# 4.3 пункт меню после пункта «Отчет по кладовщикам»; число пунктов меню +1
$pname = 'Отчет_' + $AfterReport.Split('.')[1] + '_ОткрытьОтчет'
$mi = [regex]::Match($it, [regex]::Escape($oldItem) + ',\r?\n\{8,"' + $pname + '",0,(\d+),\r?\n\{1,1,\r?\n\{"ru","[^"]*"\}\r?\n\},(\d),' + $cls + ',(\d+),1e2,0,0,1,0,(\d),0,0\}')
if (-not $mi.Success) { throw 'Interface.bin: не найден пункт меню' }
$nl = if ($mi.Value.Contains("`r`n")) { "`r`n" } else { "`n" }
$newMi = ",$($U.MenuItem),$nl{8,`"Отчет_$($Name)_ОткрытьОтчет`",0,1,$nl{1,1,$nl{`"ru`",`"$Synonym`"}$nl},1,$cls,$newId,1e2,0,0,1,0,0,0,0}"
$it = $it.Insert($mi.Index + $mi.Length, $newMi)
# начало этого меню: {5,<меню>,4,0,<число пунктов>, перед первым пунктом «Закупки»
$menuHead = [regex]::Matches($it.Substring(0, $mi.Index), '\{5,([0-9a-f-]{36}),(\d+),0,(\d+),') | Select-Object -Last 1
$cnt = [int]$menuHead.Groups[3].Value
$it = $it.Substring(0, $menuHead.Index) + $menuHead.Value.Replace(",0,$cnt,", ",0,$($cnt + 1),") + $it.Substring($menuHead.Index + $menuHead.Length)
Save "$src\Interfaces\$Interface\Ext\Interface.bin" $it
Copy-Item "$DumpRoot\Interfaces\$Interface.xml" "$src\Interfaces\$Interface.xml"
Write-Host "Интерфейс: команд $cmdCount -> $($cmdCount + 1), id пункта $newId, пунктов в меню $cnt -> $($cnt + 1)"

$files = Get-ChildItem $src -Recurse -File | ForEach-Object { $_.FullName.Substring($src.Length + 1) }
[IO.File]::WriteAllText((Join-Path $here 'files.txt'), ($files -join "`r`n"), $utf8Bom)
Write-Host "XML для частичной загрузки: $src ($($files.Count) файлов)"
if (-not $UtBase) { return }

function Invoke-Designer($argLine, $title) {
    $log = Join-Path $env:TEMP 'bl-build-config.log'
    Remove-Item $log -ErrorAction SilentlyContinue
    $p = Start-Process $Platform -ArgumentList "DESIGNER /F `"$UtBase`" /DisableStartupDialogs /DisableStartupMessages $argLine /Out `"$log`"" -PassThru
    if (-not $p.WaitForExit(40 * 60 * 1000)) { try { $p.Kill() } catch {}; throw "$title`: конфигуратор не завершился за 40 мин." }
    $out = if (Test-Path $log) { (Get-Content $log -Encoding Default) -join "`n" } else { '' }
    return [pscustomobject]@{ Code = $p.ExitCode; Out = $out }
}
$r = Invoke-Designer "/LoadConfigFromFiles `"$src`" -listFile `"$(Join-Path $here 'files.txt')`"" 'Загрузка'
Write-Host "загрузка -> $($r.Code) $($r.Out)"
if ($r.Code -ne 0) { throw 'частичная загрузка не удалась' }
$r = Invoke-Designer '/CheckModules -ThickClientOrdinaryApplication -Server -ExternalConnection' 'Проверка модулей'
$mine = ($r.Out -split "`n" | Where-Object { $_ -match $Name }) -join "`n"
Write-Host "проверка модулей -> $($r.Code); ошибок в отчете: $(if ($mine) { "`n$mine" } else { 'нет' })"
if ($mine) { throw 'ошибки компиляции в отчете' }
Write-Host 'проверка в тестовой базе завершена'
