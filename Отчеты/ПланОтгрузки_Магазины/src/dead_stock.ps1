param([string]$OutFile_)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'analyze.ps1') | Out-Null
$D = Join-Path $PSScriptRoot 'data'
$inv = [Globalization.CultureInfo]::InvariantCulture

# последнее поступление перемещением в магазин (с 01.05.2026)
$rcv = @{}
foreach ($r in (Load 'receipts.tsv')) {
    if (-not $SHOPS.ContainsKey($r.СкладКод)) { continue }
    $k = $r.СкладКод + '|' + $r.НомКод + '|' + (KeyH $r.Хар)
    if (-not $rcv.ContainsKey($k) -or $r.Дата -gt $rcv[$k]) { $rcv[$k] = $r.Дата }
}
# продажи размера по сети
$netSize = @{}
foreach ($k in $sale.Keys) { $p = $k.Split('|', 3); $nk = $p[1] + '|' + $p[2]; $netSize[$nk] = $netSize[$nk] + $sale[$k][0] }
# перемещения: донор -> количество
$mvOut = @{}
foreach ($m in $moves) { $dc = ($SHOPS.Values | Where-Object { $_.Name -eq $m.Откуда } | Select-Object -First 1).Code
    $k = $dc + '|' + $m.Код + '|' + $m.Характеристика; $mvOut[$k] = $mvOut[$k] + $m.Количество }

$rows = New-Object Collections.Generic.List[object]
foreach ($k in $stock.Keys) {
    if ($sale.ContainsKey($k) -and $sale[$k][0] -gt 0) { continue }
    $p = $k.Split('|', 3); $sc = $p[0]; $nc = $p[1]; $h = $p[2]
    $n = $nom[$nc]; $q = $stock[$k]
    $y = if ($h -match '\bY(\d\d)\b') { 'Y' + $matches[1] } else { 'без сезона' }
    $cs13 = $colorSales[$sc + '|' + $nc + '|' + (ColorOf $h)] + 0
    $last = $rcv[$k]
    $cat = if ($last -and $last -ge '2026-10-22') { '1. Поступил за последние 4 недели' }
           elseif ($y -match '^Y(\d\d)$' -and [int]$matches[1] -le 24) { '2. Старая коллекция (Y24 и раньше)' }
           elseif ($cs13 -gt 0) { '3. Модель продаётся, этот размер — нет' }
           else { '4. Модель-цвет не продаётся в магазине' }
    $pr = $price[$nc]
    $rows.Add([pscustomobject]@{
        Магазин=$SHOPS[$sc].Name; Город=$SHOPS[$sc].City; Категория=$cat; Группа=$n.Группа; Код=$nc; Артикул=$n.Артикул; Номенклатура=$n.Наименование
        Характеристика=$h; Размер=(SizeOf $h); Коллекция=$y; Остаток=$q; Цена=$pr; Сумма=$(if ($pr) { $pr * $q } else { $null })
        ПродажиМоделиВМагазине=$cs13; ПродажиРазмераПоСети=($netSize[$nc + '|' + $h] + 0); ПоследнееПоступление=$last
        ЗапрошенДругимМагазином=($mvOut[$k] + 0) })
}
$rows | Export-Csv (Join-Path $PSScriptRoot 'dead_stock.csv') -NoTypeInformation -Encoding UTF8 -Delimiter "`t"
$tot = 0.0; foreach ($r in $rows) { $tot += $r.Остаток }
"rows $($rows.Count) units $tot"
$rows | Group-Object Категория | Sort-Object Name | ForEach-Object { $u = 0.0; foreach ($r in $_.Group) { $u += $r.Остаток }; "$($_.Name): $u" }
