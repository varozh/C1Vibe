# Накладывает исправление печати QR (правки «Claude 17.11.2026») на XML-выгрузку Розницы:
#  - CommonModules\УмКо_РаботаСМаркировкой — ДополнитьКодМаркировкиРазделителямиГрупп;
#  - Documents\ПеремещениеТоваров\Forms\ФормаДокумента — модуль внутри Form.bin, кнопка «Печать QR».
# Правки вносятся точечно в текущий текст модулей; если правка уже есть, модуль не меняется.
#   apply-to-dump.ps1 -DumpRoot <каталог выгрузки>
param([Parameter(Mandatory)][string]$DumpRoot)
$ErrorActionPreference = 'Stop'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$Mark = 'Claude 17.11.2026'

function Patch([string]$text, [string]$pattern, [string]$new, [string]$what) {
    $m = [regex]::Matches($text, $pattern)
    if ($m.Count -ne 1) { throw "$what`: найдено совпадений $($m.Count) вместо 1" }
    $nl = if ($m[0].Value.Contains("`r`n")) { "`r`n" } else { "`n" }
    return $text.Substring(0, $m[0].Index) + $new.Replace("`n", $nl) + $text.Substring($m[0].Index + $m[0].Length)
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
    [IO.File]::WriteAllBytes($path, $ms.ToArray())
}
#endregion

# --- 1. Общий модуль УмКо_РаботаСМаркировкой
$cm = Join-Path $DumpRoot 'CommonModules\УмКо_РаботаСМаркировкой\Ext\Module.bsl'
$t = [IO.File]::ReadAllText($cm, [Text.Encoding]::UTF8).TrimStart([char]0xFEFF)
if ($t.Contains($Mark)) { Write-Host "УмКо_РаботаСМаркировкой: правка уже есть" } else {
    $t = Patch $t '(?m)^\t\tпозиция91 = СтрНайти\(КодМаркировки, "91",, позиция21 \+ 15\);\r?\n' (
        "`t`tпозиция91 = СтрНайти(КодМаркировки, `"91`",, позиция21 + 15);`n" +
        "`t`t// $Mark Начало`n" +
        "`t`t// Без `"91`" (короткий код или не код GS1) разделитель не вставляем:`n" +
        "`t`t// раньше при позиция91 = 0 он вставлялся в начало строки.`n" +
        "`t`tЕсли позиция91 = 0 Тогда`n`t`t`tВозврат;`n`t`tКонецЕсли;`n" +
        "`t`t// $Mark Конец`n") 'позиция91'
    $t = Patch $t '(?m)^\t\tпозиция92 = СтрНайти\(КодМаркировки, "92",, позиция91\);\r?\n' (
        "`t`t// $Mark`: `"92`" ищем после GS, `"91`" и 4 символов ключа, а не внутри ключа`n" +
        "`t`t//позиция92 = СтрНайти(КодМаркировки, `"92`",, позиция91);`n" +
        "`t`tпозиция92 = СтрНайти(КодМаркировки, `"92`",, позиция91 + 7);`n") 'позиция92'
    [IO.File]::WriteAllText($cm, $t, $utf8Bom)
    Write-Host "УмКо_РаботаСМаркировкой: правка внесена"
}

# --- 2. Модуль формы ПеремещениеТоваров.ФормаДокумента (внутри Form.bin)
$fb = Join-Path $DumpRoot 'Documents\ПеремещениеТоваров\Forms\ФормаДокумента\Ext\Form.bin'
$c = Read-Container $fb
if (-not $c.Contains('module')) { throw "Form.bin: нет файла module (есть: $($c.Keys -join ', '))" }
$t = [Text.Encoding]::UTF8.GetString($c['module'].Data).TrimStart([char]0xFEFF)
if ($t.Contains($Mark)) { Write-Host "ФормаДокумента: правка уже есть" } else {
    $t = Patch $t '(?m)^\t\tШтрихкод = СтрТ\.ПолныйКодМаркировки;\r?\n\t\tУмКо_РаботаСМаркировкой\.ДополнитьКодМаркировкиРазделителямиГрупп\(Штрихкод\);\r?\n' (
        "`t`t// $Mark Начало`n" +
        "`t`t// ПолныйКодМаркировки хранится в base64 - его передаем на печать как есть.`n" +
        "`t`t// Разделители групп вставляем только в код, хранящийся без base64 (начинается с `"01`").`n" +
        "`t`t// Раньше разделители вставлялись прямо в base64-строку и портили QR у части марок.`n" +
        "`t`t//Штрихкод = СтрТ.ПолныйКодМаркировки;`n" +
        "`t`t//УмКо_РаботаСМаркировкой.ДополнитьКодМаркировкиРазделителямиГрупп(Штрихкод);`n" +
        "`t`tШтрихкод = СокрЛП(СтрТ.ПолныйКодМаркировки);`n" +
        "`t`tЕсли Лев(Штрихкод, 2) = `"01`" Тогда`n" +
        "`t`t`tУмКо_РаботаСМаркировкой.ДополнитьКодМаркировкиРазделителямиГрупп(Штрихкод);`n" +
        "`t`t`tШтрихкод = ШтрихкодированиеИСКлиентСервер.ШтрихкодВBase64(Штрихкод);`n" +
        "`t`tКонецЕсли;`n" +
        "`t`t// $Mark Конец`n") 'Печать QR'
    $c['module'].Data = $utf8Bom.GetPreamble() + [Text.Encoding]::UTF8.GetBytes($t)
    Write-Container $fb $c
    Write-Host "ФормаДокумента: правка внесена"
}
