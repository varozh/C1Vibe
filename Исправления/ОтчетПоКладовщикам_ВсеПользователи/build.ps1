# Собирает внешний отчет «Отчет по кладовщикам (все пользователи)» (УТ 10.3) — копию Отчеты.ОтчетПоКладовщикам
# из выгрузки конфигурации, в которой ответственные ищутся среди всех пользователей, а не только в группе SKLAD.
#
#   build.ps1 [-UtBase <файловая база с конфигурацией УТ>]
#
# Изменения относительно отчета конфигурации:
#   - схема: отбор «Ответственный В ИЕРАРХИИ(&Ответственные)» заменен на «ответственный заполнен»,
#            параметр «Ответственные» удален;
#   - модуль формы: не заполняет параметр «Ответственные» (группа пользователей SKLAD);
#   - форма: тип основного реквизита — тип внешнего отчета.
param(
    [string]$UtBase,
    [string]$Platform = 'C:\Program Files (x86)\1cv8\8.3.12.1790\bin\1cv8.exe',
    [string]$DumpRoot = 'D:\002_BASE\Vibe\УТ'
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$CfgName = 'ОтчетПоКладовщикам'
$Name = 'ОтчетПоКладовщикамВсеПользователи'
$cfgRep = Join-Path $DumpRoot "Reports\$CfgName"
$src = Join-Path $here 'src'
$dst = Join-Path $src $Name
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$CR = "`r`n"

function Get-Uuid([string]$s) {
    $h = [Security.Cryptography.MD5]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes("BL-ut-storekeepers-all-users:" + $s))
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

#region Контейнер Form.bin (файлы form и module, без сжатия)
function Read-Block([byte[]]$b, [int]$addr) {
    $ms = New-Object IO.MemoryStream
    $rest = [Convert]::ToInt32([Text.Encoding]::ASCII.GetString($b, $addr + 2, 8), 16)
    while ($true) {
        $h = [Text.Encoding]::ASCII.GetString($b, $addr, 31)
        $page = [Convert]::ToInt32($h.Substring(11, 8), 16); $next = [Convert]::ToInt32($h.Substring(20, 8), 16)
        $n = [Math]::Min($page, $rest); $ms.Write($b, $addr + 31, $n); $rest -= $n
        if ($next -eq 0x7fffffff -or $rest -le 0) { break }
        $addr = $next
    }
    return , $ms.ToArray()
}
function Read-Container([string]$path) {
    $b = [IO.File]::ReadAllBytes($path); $toc = Read-Block $b 16; $res = [ordered]@{}
    for ($i = 0; $i -lt $toc.Length; $i += 12) {
        $hdr = Read-Block $b ([BitConverter]::ToInt32($toc, $i))
        $nm = [Text.Encoding]::Unicode.GetString($hdr, 20, $hdr.Length - 20).TrimEnd([char]0)
        $res[$nm] = [pscustomobject]@{ Header = $hdr; Data = (Read-Block $b ([BitConverter]::ToInt32($toc, $i + 4))) }
    }
    return $res
}
function Write-Container([string]$path, $files) {
    $ms = New-Object IO.MemoryStream
    $w = { param([byte[]]$x) $ms.Write($x, 0, $x.Length) }
    $blk = {
        param([byte[]]$data, [int]$page)
        if ($page -lt $data.Length) { $page = $data.Length }
        & $w ([Text.Encoding]::ASCII.GetBytes(("`r`n{0:x8} {1:x8} 7fffffff `r`n" -f $data.Length, $page)))
        & $w $data
        if ($page -gt $data.Length) { & $w (New-Object byte[] ($page - $data.Length)) }
    }
    & $w ([byte[]](0xFF, 0xFF, 0xFF, 0x7F, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0))
    $tocPage = 512; $addr = 16 + 31 + $tocPage; $toc = New-Object IO.MemoryStream
    foreach ($f in $files.Values) {
        $ha = $addr; $addr += 31 + $f.Header.Length; $da = $addr; $addr += 31 + $f.Data.Length
        foreach ($v in $ha, $da, 0x7fffffff) { $x = [BitConverter]::GetBytes([int]$v); $toc.Write($x, 0, 4) }
    }
    & $blk $toc.ToArray() $tocPage
    foreach ($f in $files.Values) { & $blk $f.Header 0; & $blk $f.Data 0 }
    New-Item -ItemType Directory -Force (Split-Path -Parent $path) | Out-Null
    [IO.File]::WriteAllBytes($path, $ms.ToArray())
}
#endregion

if (Test-Path $src) { Remove-Item $src -Recurse -Force }

$U = @{
    Report = Get-Uuid 'report'; ObjType = Get-Uuid 'obj-type'; ObjValue = Get-Uuid 'obj-value'
    ContObj = Get-Uuid 'contained'; Form = Get-Uuid 'form'; Template = Get-Uuid 'template'
}

#region Корневой XML внешнего отчета
$cfgXml = Read-Text "$cfgRep.xml"
$ns = [regex]::Match($cfgXml, '<MetaDataObject [^>]*>').Value
Save "$src\$Name.xml" (@"
<?xml version="1.0" encoding="UTF-8"?>
$ns
	<ExternalReport uuid="$($U.Report)">
		<InternalInfo>
			<xr:ContainedObject>
				<xr:ClassId>e41aff26-25cf-4bb6-b6c1-3f478a75f374</xr:ClassId>
				<xr:ObjectId>$($U.ContObj)</xr:ObjectId>
			</xr:ContainedObject>
			<xr:GeneratedType name="ExternalReportObject.$Name" category="Object">
				<xr:TypeId>$($U.ObjType)</xr:TypeId>
				<xr:ValueId>$($U.ObjValue)</xr:ValueId>
			</xr:GeneratedType>
		</InternalInfo>
		<Properties>
			<Name>$Name</Name>
			<Synonym>
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Отчет по кладовщикам (все пользователи)</v8:content>
				</v8:item>
			</Synonym>
			<Comment/>
			<DefaultForm>ExternalReport.$Name.Form.ФормаОтчета</DefaultForm>
			<AuxiliaryForm/>
			<MainDataCompositionSchema>ExternalReport.$Name.Template.ОсновнаяСхемаКомпоновкиДанных</MainDataCompositionSchema>
			<DefaultSettingsForm/>
			<AuxiliarySettingsForm/>
			<DefaultVariantForm/>
			<VariantsStorage/>
			<SettingsStorage/>
		</Properties>
		<ChildObjects>
			<Form>ФормаОтчета</Form>
			<Template>ОсновнаяСхемаКомпоновкиДанных</Template>
		</ChildObjects>
	</ExternalReport>
</MetaDataObject>
"@ -replace "`r?`n", $CR)
$t = Read-Text "$cfgRep\Forms\ФормаОтчета.xml"
Save "$dst\Forms\ФормаОтчета.xml" ([regex]::Replace($t, '<Form uuid="[^"]+">', "<Form uuid=`"$($U.Form)`">"))
$t = Read-Text "$cfgRep\Templates\ОсновнаяСхемаКомпоновкиДанных.xml"
Save "$dst\Templates\ОсновнаяСхемаКомпоновкиДанных.xml" ([regex]::Replace($t, '<Template uuid="[^"]+">', "<Template uuid=`"$($U.Template)`">"))
#endregion

#region Схема компоновки: без отбора по группе пользователей
$s = (Read-Text "$cfgRep\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml") -replace "`r?`n", "`n"
$s = Replace-Once $s "`tТоварыНаСкладах.Ответственный В ИЕРАРХИИ(&amp;Ответственные)" `
    "`tНЕ ТоварыНаСкладах.Ответственный ЕСТЬ NULL`n`tИ ТоварыНаСкладах.Ответственный &lt;&gt; ЗНАЧЕНИЕ(Справочник.Пользователи.ПустаяСсылка)" 'отбор по ответственным'
$s = [regex]::Replace($s, "`t<parameter>`n`t`t<name>Ответственные</name>`n(?:.*?`n)*?`t</parameter>`n", '')
$s = [regex]::Replace($s, "`t`t`t`t<dcscor:item xsi:type=`"dcsset:SettingsParameterValue`">`n`t`t`t`t`t<dcscor:parameter>Ответственные</dcscor:parameter>`n(?:.*?`n)*?`t`t`t`t</dcscor:item>`n", '')
if ($s.Contains('Ответственные')) { throw 'в схеме остался параметр Ответственные' }

# Поле «Документ» (регистратор движения) — заполняется только при параметре ПоДокументам (для расшифровки).
$doc = "ВЫБОР`n`t`tКОГДА &amp;ПоДокументам`n`t`t`tТОГДА ТоварыНаСкладах.Регистратор`n`t`tИНАЧЕ NULL`n`tКОНЕЦ"
$s = Replace-Once $s "`tКОНЕЦ КАК День,`n`tСУММА(" "`tКОНЕЦ КАК День,`n`t$doc КАК Документ,`n`tСУММА(" 'поле Документ во временной таблице'
$s = Replace-Once $s "`t`tИНАЧЕ ДАТАВРЕМЯ(1, 1, 1)`n`tКОНЕЦ`n;" "`t`tИНАЧЕ ДАТАВРЕМЯ(1, 1, 1)`n`tКОНЕЦ,`n`t$doc`n;" 'группировка по документу'
$s = Replace-Once $s "`tТоварыНаСкладах.День,`n" "`tТоварыНаСкладах.День,`n`tТоварыНаСкладах.Документ,`n" 'поле Документ в итоговом запросе'
$docField = @"
		<field xsi:type="DataSetFieldField">
			<dataPath>Документ</dataPath>
			<field>Документ</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>Документ</v8:content>
				</v8:item>
			</title>
		</field>
		<dataSource>ИсточникДанных1</dataSource>
"@ -replace "`r?`n", "`n"
$s = Replace-Once $s "`t`t<dataSource>ИсточникДанных1</dataSource>`n" $docField 'поле Документ в наборе'
$docParam = @"
	<parameter>
		<name>ПоДокументам</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>По документам</v8:content>
			</v8:item>
		</title>
		<valueType>
			<v8:Type>xs:boolean</v8:Type>
		</valueType>
		<value xsi:type="xs:boolean">false</value>
		<useRestriction>true</useRestriction>
	</parameter>
	<settingsVariant>
"@ -replace "`r?`n", "`n"
$s = Replace-Once $s "`t<settingsVariant>" $docParam 'параметр ПоДокументам'
# План и выполнение плана — и в строках складов (раньше это были детальные записи, теперь группировка «Склад»).
$n = ([regex]::Matches($s, "`t`t<group>Ответственный</group>`n")).Count
if ($n -ne 2) { throw "ожидалось 2 итога по группировке Ответственный, найдено $n" }
$s = $s.Replace("`t`t<group>Ответственный</group>`n", "`t`t<group>Ответственный</group>`n`t`t<group>Склад</group>`n")
Save "$dst\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml" ($s -replace "`n", $CR)
#endregion

#region Форма: модуль без группы SKLAD, тип основного реквизита
$c = Read-Container "$cfgRep\Forms\ФормаОтчета\Ext\Form.bin"
$f = $utf8Bom.GetString($c['form'].Data).TrimStart([char]0xFEFF) -replace "`r?`n", "`n"
$cfgType = ([xml]$cfgXml).MetaDataObject.Report.InternalInfo.GeneratedType |
    Where-Object { $_.category -eq 'Object' } | ForEach-Object { $_.TypeId }
$f = Replace-Once $f "{`"#`",$cfgType}" "{`"#`",$($U.ObjType)}" 'тип основного реквизита ОтчетОбъект'
$c['form'].Data = [byte[]]($utf8Bom.GetPreamble() + [Text.Encoding]::UTF8.GetBytes(($f -replace "`n", $CR)))

$m = $utf8Bom.GetString($c['module'].Data).TrimStart([char]0xFEFF) -replace "`r?`n", "`n"
$m = Replace-Once $m "`tОтветственные = Справочники.Пользователи.НайтиПоКоду(`"SKLAD`");`n`n`tКомпоновщикНастроек.Настройки.ПараметрыДанных.Элементы.Найти(`"Ответственные`").Значение = Ответственные;`n" `
    "`t// Ответственные по отгрузкам и приемке ищутся среди всех пользователей (в отчете конфигурации — только группа SKLAD).`n" 'заполнение параметра Ответственные'
if ($m.Contains('"Ответственные"')) { throw 'в модуле осталось обращение к параметру Ответственные' }

# Расшифровка: двойной щелчок по ячейке отчета — список документов, которые в нее вошли.
$m = "Перем НастройкиОтчета; // настройки последнего формирования — для расшифровки`n" +
     "Перем НастройкиРасшифровки Экспорт; // заполняется, если форма открыта как расшифровка по документам`n`n" + $m.TrimStart("`n")
$m = Replace-Once $m "Процедура ПриОткрытии()`n`n" `
    ("Процедура ПриОткрытии()`n`n`tЕсли НастройкиРасшифровки <> Неопределено Тогда`n" +
     "`t`tКомпоновщикНастроек.ЗагрузитьНастройки(НастройкиРасшифровки);`n`tКонецЕсли;`n`n") 'настройки расшифровки при открытии'
$m = Replace-Once $m "`tНастройкиКомпоновкиДанных = КомпоновщикНастроек.ПолучитьНастройки();`n" `
    ("`tСхемаКомпоновкиДанных.Параметры.Найти(`"ПоДокументам`").Значение = (НастройкиРасшифровки <> Неопределено);`n`n" +
     "`tНастройкиКомпоновкиДанных = КомпоновщикНастроек.ПолучитьНастройки();`n" +
     "`tНастройкиОтчета = КомпоновщикНастроек.ПолучитьНастройки();`n") 'запоминание настроек'
$m = Replace-Once $m "`tЕсли ПоДням Тогда`n`t`tНастроитьВыводПоДням(НастройкиКомпоновкиДанных);`n`tКонецЕсли;`n" `
    "`tЕсли ПоДням Тогда`n`t`tНастроитьВыводПоДням(НастройкиКомпоновкиДанных);`n`tИначе`n`t`tЗаменитьДетальныеЗаписиНаСклад(НастройкиКомпоновкиДанных.Структура);`n`tКонецЕсли;`n" 'группировка по складу вместо детальных записей'
$m = $m.TrimEnd() + "`n`n" + ((Read-Text (Join-Path $here 'Расшифровка.bsl')).TrimStart([char]0xFEFF) -replace "`r?`n", "`n").Trim() + "`n"
$evt = @"
{1,
{1,e1692cc2-605b-4535-84dd-28440238746c,
{3,"РезультатОбработкаРасшифровки",
{1,"РезультатОбработкаРасшифровки",
{1,1,
{"ru","Результат обработка расшифровки"}
},
{1,1,
{"ru","Результат обработка расшифровки"}
},
{1,1,
{"ru","Результат обработка расшифровки"}
},
{4,0,
{0},"",-1,-1,1,0,""},
{0,0,0}
}
}
}
},1,0,1,0,0,1,0,1,1,0,0,0,0,1,1},
"@ -replace "`r?`n", "`n"
$f = Replace-Once $f "`n{0},1,0,1,0,0,1,0,1,1,0,0,0,0,1,1},`n" ("`n" + $evt + "`n") 'событие ОбработкаРасшифровки поля Результат'
$c['form'].Data = [byte[]]($utf8Bom.GetPreamble() + [Text.Encoding]::UTF8.GetBytes(($f -replace "`n", $CR)))
Save (Join-Path $here 'ОтчетПоКладовщикамВсеПользователи_МодульФормы.bsl') ($m.TrimStart("`n") -replace "`n", $CR)
$c['module'].Data = [byte[]]($utf8Bom.GetPreamble() + [Text.Encoding]::UTF8.GetBytes(($m -replace "`n", $CR)))
Write-Container "$dst\Forms\ФормаОтчета\Ext\Form.bin" $c
#endregion

Write-Host "Исходники: $src"
if (-not $UtBase) { return }

$erf = Join-Path $here "$Name.erf"
$log = Join-Path $env:TEMP 'bl-build-storekeepers-all.log'
Remove-Item $log, $erf -ErrorAction SilentlyContinue
$p = Start-Process $Platform -ArgumentList "DESIGNER /F `"$UtBase`" /DisableStartupDialogs /DisableStartupMessages /LoadExternalDataProcessorOrReportFromFiles `"$src\$Name.xml`" `"$erf`" /Out `"$log`"" -PassThru -Wait
$out = if (Test-Path $log) { (Get-Content $log -Encoding Default) -join "`n" } else { '' }
Write-Host "сборка -> $($p.ExitCode) $out"
if ($p.ExitCode -ne 0 -or -not (Test-Path $erf)) { throw 'не удалось собрать .erf' }
Write-Host "готово: $erf"
