# Собирает обработку «Инициализировать локальный модуль» (ПИОТ) с полем «Токен» на форме.
#   build.ps1 -RetailBase <файловая база с конфигурацией Розницы> [-Out <путь .epf>]
# Исходник — ИнициализироватьЛокальныйМодуль_исходная.epf рядом со скриптом.
param(
    [Parameter(Mandatory)][string]$RetailBase,
    [string]$Out,
    [string]$Platform = 'C:\Program Files (x86)\1cv8\8.3.12.1790\bin\1cv8.exe'
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Out) { $Out = Join-Path $here 'ИнициализироватьЛокальныйМодуль.epf' }
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$CR = "`r`n"
$work = Join-Path ([IO.Path]::GetTempPath()) ('piot-token-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $work | Out-Null

function Read-Text([string]$path) { [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) }
function Replace-Once([string]$text, [string]$old, [string]$new, [string]$what) {
    $i = $text.IndexOf($old, [StringComparison]::Ordinal)
    if ($i -lt 0) { throw "Не найдено: $what" }
    if ($text.IndexOf($old, $i + 1, [StringComparison]::Ordinal) -ge 0) { throw "Найдено больше одного: $what" }
    return $text.Substring(0, $i) + $new + $text.Substring($i + $old.Length)
}
function Get-Between([string]$text, [string]$start, [string]$end, [string]$what) {
    $i = $text.IndexOf($start, [StringComparison]::Ordinal); if ($i -lt 0) { throw "Не найдено начало: $what" }
    $j = $text.IndexOf($end, $i, [StringComparison]::Ordinal); if ($j -lt 0) { throw "Не найден конец: $what" }
    return $text.Substring($i, $j + $end.Length - $i)
}
function Invoke-Designer([string]$arguments, [string]$log) {
    $p = Start-Process $Platform -ArgumentList "DESIGNER /F `"$RetailBase`" /DisableStartupDialogs /DisableStartupMessages $arguments /Out `"$log`"" -PassThru -Wait
    $text = if (Test-Path $log) { Get-Content $log -Raw } else { '' }
    if ($p.ExitCode -ne 0) { throw "Конфигуратор вернул $($p.ExitCode): $text" }
    $text
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

# 1. Выгрузка исходной обработки в XML
$src = Join-Path $work 'src'
Invoke-Designer "/DumpExternalDataProcessorOrReportToFiles `"$src\root.xml`" `"$here\ИнициализироватьЛокальныйМодуль_исходная.epf`"" "$work\dump.log" | Out-Null
$bin = "$src\root\Forms\Форма\Ext\Form.bin"
$c = Read-Container $bin
$f = $utf8Bom.GetString($c['form'].Data).TrimStart([char]0xFEFF) -replace "`r?`n", "`n"
if ($f -notmatch '"Магазин",\n\{"Pattern",\n\{"#",') { throw 'Реквизит Магазин выгружен не ссылочным типом: база без конфигурации Розницы?' }

# 2. Форма: надпись «Токен:» (id 5) и поле Токен (id 6) под магазином, реквизит Токен (id 4, строка)
$label = Get-Between $f '{0fc7e20d-f241-460c-bdf4-5ad88e5474a5,3,' "{14,`"НадписьМагазин`",4294967295,0,0,0},`n{0}`n}" 'надпись магазина'
$newLabel = $label.Replace('{0fc7e20d-f241-460c-bdf4-5ad88e5474a5,3,', '{0fc7e20d-f241-460c-bdf4-5ad88e5474a5,5,').
    Replace('{"ru","Магазин:"}', '{"ru","Токен:"}').Replace('{8,8,8,58,28,1,', '{8,8,36,58,56,1,').
    Replace('{2,3,0,20}', '{2,5,0,20}').Replace('{2,3,2,50}', '{2,5,2,50}').
    Replace('{0,3,1}', '{0,5,1}').Replace('{0,3,3}', '{0,5,3}').Replace('"НадписьМагазин"', '"НадписьТокен"')
$field = Get-Between $f '{381ed624-9217-4e63-85db-c4c3cb87daae,4,' "{14,`"Магазин`",4294967295,0,0,0},`n{0}`n}" 'поле магазина'
# Поле строкового типа (образец — то же поле, выгруженное как строка)
$strType = "{`"Pattern`",`n{`"S`"}`n}"
$newField = [regex]::Replace($field, '^\{381ed624-9217-4e63-85db-c4c3cb87daae,4,\n\{9,\n\{"Pattern",\n\{"#",[0-9a-f-]+\}\n\}', "{381ed624-9217-4e63-85db-c4c3cb87daae,6,`n{9,`n$strType")
$newField = [regex]::Replace($newField, '\{8585207b-9ca7-425b-9385-e8fd67c4e148,\n\{5,0,2,0{8}-0{4}-0{4}-0{4}-0{12},\n\{1,0\}\n\}',
    "{9a7643d2-19e9-45e2-8893-280bc9195a97,`n{4,`n{`"U`"},`n{`"U`"},0,`"`",0,0}")
$newField = $newField.Replace('16777215,1}', '16777215,2}').Replace('{8,63,8,263,28,1,', '{8,63,36,392,56,1,').
    Replace('{2,4,0,20}', '{2,6,0,20}').Replace('{2,0,3,-137}', '{2,0,3,0}').Replace('{0,4,1}', '{0,6,1}').
    Replace('"Магазин"', '"Токен"')
if ($newField -match '"#",' -or $newField -notmatch '9a7643d2') { throw 'Не удалось сделать поле строковым' }
$fieldEnd = "{14,`"Магазин`",4294967295,0,0,0},`n{0}`n}"
$f = Replace-Once $f $fieldEnd ($fieldEnd + ",`n" + $newLabel + ",`n" + $newField) 'новые элементы после поля магазина'
$f = Replace-Once $f "{3,`n{e69bf21d-97b2-4f37-86db-675aea9ec2cb,2," "{5,`n{e69bf21d-97b2-4f37-86db-675aea9ec2cb,2," 'число элементов формы'
$attrs = Get-Between $f "{`n{1},4,`n{2," "{`n{2},0,0,1,`"Магазин`"," 'список реквизитов'
$f = Replace-Once $f $attrs ($attrs.Replace("{`n{1},4,`n{2,", "{`n{1},5,`n{3,")) 'счетчик реквизитов'
$attrShop = Get-Between $f "{`n{2},0,0,1,`"Магазин`"," "}`n}`n}" 'реквизит Магазин'
$f = Replace-Once $f $attrShop ($attrShop + ",`n{`n{4},0,0,1,`"Токен`",`n$strType`n}") 'реквизит Токен'
$f = Replace-Once $f "{1,`n{4,`n{1,`n{2}`n}`n}`n}" "{2,`n{4,`n{1,`n{2}`n}`n},`n{6,`n{1,`n{4}`n}`n}`n}" 'привязка поля к реквизиту'
$c['form'].Data = [byte[]]($utf8Bom.GetPreamble() + [Text.Encoding]::UTF8.GetBytes(($f -replace "`n", $CR)))

# 3. Модуль формы
$m = (Read-Text (Join-Path $here 'ИнициализироватьЛокальныйМодуль_МодульФормы.bsl')).TrimStart([char]0xFEFF) -replace "`r?`n", $CR
$c['module'].Data = [byte[]]($utf8Bom.GetPreamble() + [Text.Encoding]::UTF8.GetBytes($m))
Write-Container $bin $c

# 4. Сборка .epf
Invoke-Designer "/LoadExternalDataProcessorOrReportFromFiles `"$src\root.xml`" `"$Out`"" "$work\load.log"
Remove-Item $work -Recurse -Force
"Собрано: $Out"
