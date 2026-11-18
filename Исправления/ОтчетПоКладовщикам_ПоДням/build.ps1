# Собирает исходники внешнего отчета «Отчет по кладовщикам» (УТ 10.3) с флажком «По дням»
# из выгрузки конфигурации (Отчеты.ОтчетПоКладовщикам) и, если указана -UtBase, сам .erf.
#
#   build.ps1 [-UtBase <файловая база с конфигурацией УТ>]
#
# Изменения относительно отчета конфигурации:
#   - схема: поле «День» (начало дня движения; заполняется только при параметре ПоДням), параметр ПоДням;
#   - форма: реквизит и флажок «По дням» справа от периода;
#   - модуль формы: при флажке — таблица с колонками по дням, итог за период первой и последней колонкой.
param(
    [string]$UtBase,
    [string]$Platform = 'C:\Program Files (x86)\1cv8\8.3.12.1790\bin\1cv8.exe',
    [string]$DumpRoot = 'D:\002_BASE\Vibe\УТ'
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Name = 'ОтчетПоКладовщикам'
$cfgRep = Join-Path $DumpRoot "Reports\$Name"
$src = Join-Path $here 'src'
$dst = Join-Path $src $Name
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$CR = "`r`n"

function Get-Uuid([string]$s) {
    $h = [Security.Cryptography.MD5]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes("BL-ut-storekeepers-by-day:" + $s))
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
					<v8:content>Отчет по кладовщикам</v8:content>
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

#region Схема компоновки: поле «День», параметр ПоДням
$s = Read-Text "$cfgRep\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml"
$s = $s -replace "`r?`n", "`n"
$day = "ВЫБОР`n`t`tКОГДА &amp;ПоДням`n`t`t`tТОГДА НАЧАЛОПЕРИОДА(ТоварыНаСкладах.Период, ДЕНЬ)`n`t`tИНАЧЕ ДАТАВРЕМЯ(1, 1, 1)`n`tКОНЕЦ"
$s = Replace-Once $s "`tЗНАЧЕНИЕ(Справочник.Пользователи.ПустаяСсылка) КАК ПустойПользователь,`n`tСУММА(" `
    "`tЗНАЧЕНИЕ(Справочник.Пользователи.ПустаяСсылка) КАК ПустойПользователь,`n`t$day КАК День,`n`tСУММА(" 'поле День во временной таблице'
$s = Replace-Once $s "`t`tИНАЧЕ ТоварыНаСкладах.Регистратор.Ответственный`n`tКОНЕЦ`n;" `
    "`t`tИНАЧЕ ТоварыНаСкладах.Регистратор.Ответственный`n`tКОНЕЦ,`n`t$day`n;" 'группировка по дню'
$s = Replace-Once $s "`tТоварыНаСкладах.ПустойПользователь,`n`tТоварыНаСкладах.КоличествоПриход," `
    "`tТоварыНаСкладах.ПустойПользователь,`n`tТоварыНаСкладах.День,`n`tТоварыНаСкладах.КоличествоПриход," 'поле День в итоговом запросе'
$dayField = @"
		<field xsi:type="DataSetFieldField">
			<dataPath>День</dataPath>
			<field>День</field>
			<title xsi:type="v8:LocalStringType">
				<v8:item>
					<v8:lang>ru</v8:lang>
					<v8:content>День</v8:content>
				</v8:item>
			</title>
			<role>
				<dcscom:periodNumber>1</dcscom:periodNumber>
				<dcscom:periodType>Main</dcscom:periodType>
			</role>
			<appearance>
				<dcscor:item xsi:type="dcsset:SettingsParameterValue">
					<dcscor:parameter>Формат</dcscor:parameter>
					<dcscor:value xsi:type="v8:LocalStringType">
						<v8:item>
							<v8:lang>ru</v8:lang>
							<v8:content>ДФ=dd.MM.yyyy</v8:content>
						</v8:item>
					</dcscor:value>
				</dcscor:item>
			</appearance>
		</field>

"@ -replace "`r?`n", "`n"
$s = Replace-Once $s "`t`t<dataSource>ИсточникДанных1</dataSource>`n`t`t<query>" ($dayField.TrimEnd("`n") + "`n`t`t<dataSource>ИсточникДанных1</dataSource>`n`t`t<query>") 'поле День в наборе'
$dayParam = @"
	<parameter>
		<name>ПоДням</name>
		<title xsi:type="v8:LocalStringType">
			<v8:item>
				<v8:lang>ru</v8:lang>
				<v8:content>По дням</v8:content>
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
$s = Replace-Once $s "`t<settingsVariant>" $dayParam 'параметр ПоДням'
Save "$dst\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml" ($s -replace "`n", $CR)
#endregion

#region Форма: реквизит ПоДням и флажок «По дням»; модуль
$c = Read-Container "$cfgRep\Forms\ФормаОтчета\Ext\Form.bin"
$f = $utf8Bom.GetString($c['form'].Data).TrimStart([char]0xFEFF) -replace "`r?`n", "`n"
$cfgType = ([xml](Read-Text "$cfgRep.xml")).MetaDataObject.Report.InternalInfo.GeneratedType |
    Where-Object { $_.category -eq 'Object' } | ForEach-Object { $_.TypeId }
# 7 элементов на панели -> 8
$f = Replace-Once $f "{7,`n{236a17b3-7f44-46d9-a907-75f9cdc61ab5,1," "{8,`n{236a17b3-7f44-46d9-a907-75f9cdc61ab5,1," 'число элементов формы'
# Флажок (образец оформления — флажок формы обработки «Отчет по данным контрагента», без привязок к соседям)
$checkbox = @"
{35af3d93-d7c7-4a2e-a8eb-bac87a1a3f26,13,
{1,
{
{
{10,1,
{3,4,
{0}
},
{3,4,
{0}
},
{6,3,0,1},0,
{3,3,
{-22}
},
{3,4,
{0}
},
{3,4,
{0}
},
{3,3,
{-7}
},
{3,3,
{-21}
},
{3,0,
{0},0,0,0,48312c09-257f-4b29-b280-284dd89efc1e},
{1,0}
},4,
{1,1,
{"ru","По дням"}
},1,0,1},1,0,0},
{0}
},
{8,336,32,426,52,1,
{0,
{2,-1,6,0},
{2,-1,6,0}
},
{0,
{2,13,0,20},
{2,-1,6,0}
},
{0,
{2,-1,6,0},
{2,-1,6,0}
},
{0,
{2,13,2,90},
{2,-1,6,0}
},
{0,
{2,-1,6,0},
{2,-1,6,0}
},
{0,
{2,-1,6,0},
{2,-1,6,0}
},1,
{0,13,1},0,1,
{0,13,3},0,0,0,0,7,7,0,0},
{14,"ПоДням",4294967295,0,0,0},
{0}
}
"@ -replace "`r?`n", "`n"
$last = "{14,`"ВыбратьПериод`",4294967295,0,0,0},`n{0}`n}"
$f = Replace-Once $f $last ($last + ",`n" + $checkbox) 'флажок после кнопки выбора периода'
$f = Replace-Once $f "{`n{1},11,`n{3," "{`n{1},12,`n{4," 'список реквизитов формы'
$attrKon = "{`n{8},1,0,1,`"ДатаКон`",`n{`"Pattern`",`n{`"D`"}`n}`n}"
$f = Replace-Once $f $attrKon ($attrKon + ",`n{`n{11},1,0,1,`"ПоДням`",`n{`"Pattern`",`n{`"B`"}`n}`n}") 'реквизит ПоДням'
$f = Replace-Once $f "{2,`n{7,`n{1,`n{7}`n}`n},`n{9,`n{1,`n{8}`n}`n}" `
    "{3,`n{7,`n{1,`n{7}`n}`n},`n{9,`n{1,`n{8}`n}`n},`n{13,`n{1,`n{11}`n}`n}" 'привязка флажка к реквизиту'
$m = (Read-Text (Join-Path $here 'ОтчетПоКладовщикам_МодульФормы.bsl')).TrimStart([char]0xFEFF) -replace "`r?`n", $CR
$c['module'].Data = [byte[]]($utf8Bom.GetPreamble() + [Text.Encoding]::UTF8.GetBytes($CR + $m))

# Вариант для конфигурации (config\src): тот же отчет с прежними uuid и типом ОтчетОбъект.ОтчетПоКладовщикам
$cfgSrc = Join-Path $here 'config\src'
$cfgDst = Join-Path $cfgSrc 'Reports'
if (Test-Path $cfgSrc) { Remove-Item $cfgSrc -Recurse -Force }
New-Item -ItemType Directory -Force $cfgDst | Out-Null
Copy-Item "$cfgRep.xml" $cfgDst
Copy-Item $cfgRep $cfgDst -Recurse
Copy-Item "$dst\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml" "$cfgDst\$Name\Templates\ОсновнаяСхемаКомпоновкиДанных\Ext\Template.xml" -Force
$c['form'].Data = [byte[]]($utf8Bom.GetPreamble() + [Text.Encoding]::UTF8.GetBytes(($f -replace "`n", $CR)))
Write-Container "$cfgDst\$Name\Forms\ФормаОтчета\Ext\Form.bin" $c

# Внешний отчет: тип основного реквизита формы — тип внешнего отчета
$fx = Replace-Once $f "{`"#`",$cfgType}" "{`"#`",$($U.ObjType)}" 'тип основного реквизита ОтчетОбъект'
$c['form'].Data = [byte[]]($utf8Bom.GetPreamble() + [Text.Encoding]::UTF8.GetBytes(($fx -replace "`n", $CR)))
Write-Container "$dst\Forms\ФормаОтчета\Ext\Form.bin" $c
#endregion

Write-Host "Исходники: $src"
if (-not $UtBase) { return }

$erf = Join-Path $here "$Name.erf"
$log = Join-Path $env:TEMP 'bl-build-storekeepers.log'
Remove-Item $log, $erf -ErrorAction SilentlyContinue
$p = Start-Process $Platform -ArgumentList "DESIGNER /F `"$UtBase`" /DisableStartupDialogs /DisableStartupMessages /LoadExternalDataProcessorOrReportFromFiles `"$src\$Name.xml`" `"$erf`" /Out `"$log`"" -PassThru -Wait
$out = if (Test-Path $log) { (Get-Content $log -Encoding Default) -join "`n" } else { '' }
Write-Host "сборка -> $($p.ExitCode) $out"
if ($p.ExitCode -ne 0 -or -not (Test-Path $erf)) { throw 'не удалось собрать .erf' }
Write-Host "готово: $erf"
