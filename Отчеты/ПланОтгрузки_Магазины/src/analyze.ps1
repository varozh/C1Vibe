param(
    [double]$Season = 1.2,      # рост продаж в декабре к окт-ноя (по прошлому году)
    [double]$Cycle = 1,         # период между отгрузками, недель
    [int]$NewTopPerStore = 12,  # новинок (модель-цвет) на магазин
    [int]$NewMinNet = 8         # мин. продаж новинки по сети за 4 недели
)
$ErrorActionPreference = 'Stop'
$D = Join-Path $PSScriptRoot 'data'
$inv = [Globalization.CultureInfo]::InvariantCulture
function Num($s) { if (-not $s) { return 0.0 } [double]::Parse($s, $inv) }
function Load($f) { Import-Csv (Join-Path $D $f) -Delimiter "`t" -Encoding UTF8 }
function KeyH([string]$h) { $h = $h.Trim(); if ($h.Length -gt 100) { $h = $h.Substring(0,100) }; $h.TrimEnd() }
function SizeOf([string]$h) { if ($h -match 'Размер ([^,]+)') { $matches[1].Trim() } else { '' } }
function ColorOf([string]$h) { ($h -replace ',\s*Размер [^,]+', '').Trim() }

# Магазины: код склада, магазин, город, срок доставки (недель)
$stores = @(
 @('ЦБР000597','Влд ТЦ Мегаторг','Владимир',1), @('ЦБР000691','Влг ТЦ Мармелад','Волгоград',1),
 @('ЦБР000541','Екб ТРК ВеерМолл','Екатеринбург',2), @('ЦБР000536','Екб ТРК Гринвич','Екатеринбург',2), @('ЦБР000557','Екб ТРК Радуга','Екатеринбург',2),
 @('ЦБР000599','Кзн ТРЦ Парк Хаус','Казань',1), @('ЦБР000120','Кзн ТЦ Кольцо','Казань',1),
 @('ЦБР000523','Мск ТЦ Афимолл','Москва',1), @('ЦБР000330','Мск ТЦ Ереван Плаза','Москва',1), @('ЦБР000041','Мск ТЦ Таганский Пассаж','Москва',1),
 @('ЦБР000089','НН ТЦ Седьмое небо','Нижний Новгород',1), @('ЦБР000091','НН ТЦ Фантастика','Нижний Новгород',1),
 @('ЦБР000695','Орб ТЦ Восход','Оренбург',1), @('ЦБР000676','Орб ТЦ Мармелад','Оренбург',1),
 @('ЦБР000090','Пнз ТЦ Коллаж','Пенза',1), @('ЦБР000069','Пнз ТЦ САНиМАРТ','Пенза',1),
 @('ЦБР000566','Прм ТРК Планета','Пермь',2),
 @('ЦБР000115','Смр ТЦ Аврора','Самара',1), @('ЦБР000462','Смр ТЦ Амбар','Самара',1), @('ЦБР000101','Смр ТЦ Вива Лэнд','Самара',1),
 @('ЦБР000675','СП ТРК Охта Молл','Санкт-Петербург',2), @('ЦБР000592','СП ТРК Французский Бульвар','Санкт-Петербург',2),
 @('ЦБР000135','Саранск ТЦ Сити Парк','Саранск',1),
 @('ЦБР000286','Саратов ТЦ HAPPY Молл','Саратов',1), @('ЦБР000364','Саратов ТЦ Оранжевый','Саратов',1), @('ЦБР000296','Саратов ТЦ Форум','Саратов',1),
 @('ЦБР000533','Сочи МореМолл','Сочи',2),
 @('ЦБР000039','Улн ТЦ Аквамолл','Ульяновск',1), @('ЦБР000046','Улн ТЦ Самолет','Ульяновск',1), @('ЦБР000305','ОФИС (E-COM)','Ульяновск',1),
 @('ЦБР000672','Уфа ТЦ Мир','Уфа',1),
 @('ЦБР000104','Чбк ТЦ Мегамолл','Чебоксары',1),
 @('ЦБР000594','Члб ТЦ Алмаз','Челябинск',2), @('ЦБР000690','Члб ТЦ Родник','Челябинск',2)
)
$SHOPS = @{}; foreach ($x in $stores) { $SHOPS[$x[0]] = [pscustomobject]@{ Code=$x[0]; Name=$x[1]; City=$x[2]; Lead=[double]$x[3] } }

# Номенклатура
$nom = @{}
foreach ($r in (Load 'nom.tsv')) { $nom[$r.Код] = $r }
function GroupOf($code) { $n = $nom[$code]; if (-not $n) { return '' } $n.Группа }
function IsGoods($code) {
    $g = GroupOf $code
    $g2 = $nom[$code].Группа2
    if ($g2 -match '^(7|8|9)\.' -or $g -match 'ВЕШАЛК|POS-|Видеонаблюд|Сотовые|РАЗВЕС|ТОРГОВОЕ') { return $false }
    -not ($g -match '^(7|8|9)\.' -or $g -match 'СЕРТИФИКАТ' -or $g -match 'ОБОРУДОВАН' -or $g -eq '')
}

# Продажи
$weeks = @{}
$rows = @(Load 'sales.tsv') + @(Load 'sales_ofis.tsv')
$sale = @{}   # store|nom|har -> @(s13, s4)
$firstWeek = @{}
$net4 = @{}   # nom|color -> network sales 4w
$storeSales = @{}
foreach ($r in $rows) {
    if (-not $SHOPS.ContainsKey($r.СкладКод)) { continue }
    $q = Num $r.Кол
    $w = $r.Неделя
    if (-not $firstWeek.ContainsKey($r.СкладКод) -or $w -lt $firstWeek[$r.СкладКод]) { $firstWeek[$r.СкладКод] = $w }
    if (-not (IsGoods $r.НомКод)) { continue }
    $k = $r.СкладКод + '|' + $r.НомКод + '|' + (KeyH $r.Хар)
    if (-not $sale.ContainsKey($k)) { $sale[$k] = @(0.0, 0.0, 0.0) }
    $sale[$k][0] += $q; $sale[$k][2] += (Num $r.Сумма)
    if ($w -ge '2026-10-26') {
        $sale[$k][1] += $q
        $ck = $r.НомКод + '|' + (ColorOf (KeyH $r.Хар))
        $net4[$ck] = $net4[$ck] + $q
    }
    $storeSales[$r.СкладКод] = $storeSales[$r.СкладКод] + $q
}
$price = @{}; $pq = @{}
foreach ($k in $sale.Keys) { $nc = $k.Split('|',3)[1]; if ($sale[$k][2] -gt 0) { $price[$nc] = $price[$nc] + $sale[$k][2]; $pq[$nc] = $pq[$nc] + $sale[$k][0] } }
foreach ($nc in @($price.Keys)) { $price[$nc] = [Math]::Round($price[$nc] / $pq[$nc]) }
# Остатки магазинов
$stock = @{}; $storeColor = @{}; $storeStock = @{}
foreach ($r in (@(Load 'stock.tsv') + @(Load 'stock_ofis.tsv'))) {
    if (-not $SHOPS.ContainsKey($r.СкладКод)) { continue }
    if (-not (IsGoods $r.НомКод)) { continue }
    $q = Num $r.Кол; if ($q -le 0) { continue }
    $h = KeyH $r.Хар
    $k = $r.СкладКод + '|' + $r.НомКод + '|' + $h
    $stock[$k] = $stock[$k] + $q
    $storeColor[$r.СкладКод + '|' + $r.НомКод + '|' + (ColorOf $h)] = 1
    $storeStock[$r.СкладКод] = $storeStock[$r.СкладКод] + $q
}
# В пути (перемещения с 20.10, не принятые)
$transit = @{}
foreach ($r in (Load 'transit.tsv')) {
    if (-not $SHOPS.ContainsKey($r.СкладКод)) { continue }
    $h = KeyH $r.Хар
    $k = $r.СкладКод + '|' + $r.НомКод + '|' + $h
    $transit[$k] = $transit[$k] + (Num $r.Кол)
    $storeColor[$r.СкладКод + '|' + $r.НомКод + '|' + (ColorOf $h)] = 1
}
# Центральный склад (УТ, Главный розничный склад)
$cs = @{}; $csFull = @{}
foreach ($r in (Load 'grs.tsv')) {
    if ($r.СкладКод -ne 'УТ0000389' -or $r.Кач -ne 'Новый') { continue }
    $q = Num $r.Кол; if ($q -le 0) { continue }
    $k = $r.НомКод + '|' + (KeyH $r.Хар)
    $cs[$k] = $cs[$k] + $q; $csFull[$k] = $r.Хар
}
$csStart = @{} + $cs

# Активные недели магазина
$end = [datetime]'2026-11-19'
function Weeks($code, [double]$cap) {
    $f = $firstWeek[$code]; if (-not $f) { return $cap }
    $w = (($end - [datetime]$f).TotalDays + 1) / 7.0
    [Math]::Min($cap, [Math]::Max(1.0, $w))
}

# 1. Пополнение проданного
$lines = New-Object Collections.Generic.List[object]
foreach ($k in $sale.Keys) {
    $p = $k.Split('|', 3); $sc = $p[0]; $nc = $p[1]; $h = $p[2]
    $s13 = $sale[$k][0]; $s4 = $sale[$k][1]
    if ($s13 -le 0 -or $sale[$k][2] -le 0) { continue }
    $W13 = Weeks $sc 13.57; $W4 = Weeks $sc 3.57
    $rate = 0.5 * $s4 / $W4 + 0.5 * $s13 / $W13
    $st = $stock[$k] + 0; $tr = $transit[$k] + 0
    $L = $SHOPS[$sc].Lead
    $cover = $L + $Cycle
    $target = [Math]::Ceiling($rate * $Season * $cover - 0.25)
    if ($target -lt 1) { $target = 1 }   # продаваемый размер держим в наличии
    $need = $target - $st - $tr
    if ($need -le 0) { continue }
    $lines.Add([pscustomobject]@{ Store=$sc; Nom=$nc; Har=$h; Type='Пополнение'; S13=$s13; S4=$s4; Rate=[Math]::Round($rate,2); Stock=$st; Transit=$tr; Target=$target; Need=$need; Qty=0; Prio = if ($st + $tr -eq 0) { 0 } else { 1 } })
}
# распределение пополнения по остатку центрального склада: сначала нулевые остатки, затем по скорости; дальние магазины раньше при равенстве
$sorted = $lines | Sort-Object Prio, @{e={-$_.Rate}}, @{e={-$SHOPS[$_.Store].Lead}}
foreach ($l in $sorted) {
    $ck = $l.Nom + '|' + $l.Har
    $av = $cs[$ck] + 0
    if ($av -le 0) { continue }
    $q = [Math]::Min($av, $l.Need); $l.Qty = $q; $cs[$ck] = $av - $q
}

# 2. Новинки: модель-цвет, хорошо продающиеся по сети, которых нет в магазине
$csColor = @{}   # nom|color -> list of (harKey)
foreach ($k in $cs.Keys) {
    if ($cs[$k] -le 0) { continue }
    $p = $k.Split('|', 2); if (-not (IsGoods $p[0])) { continue }
    $ck = $p[0] + '|' + (ColorOf $p[1])
    if (-not $csColor.ContainsKey($ck)) { $csColor[$ck] = New-Object Collections.Generic.List[string] }
    $csColor[$ck].Add($p[1])
}
$cand = $net4.Keys | Where-Object { $net4[$_] -ge $NewMinNet -and $csColor.ContainsKey($_) } | Sort-Object @{e={-$net4[$_]}}
# магазины по убыванию продаж
$storeOrder = $SHOPS.Keys | Sort-Object @{e={-($storeSales[$_] + 0)}}
$newCount = @{}
foreach ($ck in $cand) {
    $p = $ck.Split('|', 2)
    foreach ($sc in $storeOrder) {
        if ($sc -eq 'ЦБР000305') { continue }
        if ($storeColor.ContainsKey($sc + '|' + $ck)) { continue }
        if (($newCount[$sc] + 0) -ge $NewTopPerStore) { continue }
        # размерный ряд: размеры, доступные на складе (не более 1 шт каждого)
        $sizes = $csColor[$ck] | Where-Object { $cs[$p[0] + '|' + $_] -gt 0 }
        $sizes = @($sizes | Where-Object { $z = SizeOf $_; $z -notmatch '^(38|40|56|58|60)$' })
        if ($sizes.Count -lt 2) { continue }
        foreach ($h in $sizes) {
            $kk = $p[0] + '|' + $h
            $lines.Add([pscustomobject]@{ Store=$sc; Nom=$p[0]; Har=$h; Type='Новинка'; S13=0; S4=0; Rate=0; Stock=0; Transit=0; Target=1; Need=1; Qty=1; Prio=2 })
            $cs[$kk] = $cs[$kk] - 1
        }
        $newCount[$sc] = $newCount[$sc] + 1
    }
}

# 3. Перемещения между магазинами под дефицит ЦС: донор — магазин, где размер лежит без продаж 13 недель
$colorSales = @{}
foreach ($k in $sale.Keys) { $p = $k.Split('|', 3); $ck = $p[0] + '|' + $p[1] + '|' + (ColorOf $p[2]); $colorSales[$ck] = $colorSales[$ck] + $sale[$k][0] }
$donorAv = @{}   # nom|har -> list of donor store codes
foreach ($k in $stock.Keys) {
    $p = $k.Split('|', 3)
    if ($p[0] -eq 'ЦБР000305') { }
    if (($sale[$k] -ne $null) -and $sale[$k][0] -gt 0) { continue }
    if (($colorSales[$p[0] + '|' + $p[1] + '|' + (ColorOf $p[2])] + 0) -gt 1) { continue }
    $nk = $p[1] + '|' + $p[2]
    if (-not $donorAv.ContainsKey($nk)) { $donorAv[$nk] = New-Object Collections.Generic.List[string] }
    $donorAv[$nk].Add($p[0])
}
$moves = New-Object Collections.Generic.List[object]
$donorLeft = @{}; foreach ($k in $stock.Keys) { $donorLeft[$k] = $stock[$k] }
foreach ($l in ($lines | Where-Object { $_.Type -eq 'Пополнение' -and $_.Need - $_.Qty -gt 0 } | Sort-Object Prio, @{e={-$_.Rate}})) {
    $nk = $l.Nom + '|' + $l.Har
    if (-not $donorAv.ContainsKey($nk)) { continue }
    $rem = $l.Need - $l.Qty
    $city = $SHOPS[$l.Store].City
    $cands = $donorAv[$nk] | Where-Object { $_ -ne $l.Store } | Sort-Object @{e={ if ($SHOPS[$_].City -eq $city) { 0 } else { 1 } }}, @{e={ -$donorLeft[$_ + '|' + $nk] }}
    foreach ($d in $cands) {
        if ($rem -le 0) { break }
        $dk = $d + '|' + $nk
        $left = $donorLeft[$dk] + 0
        if ($left -le 0) { continue }
        $q = [Math]::Min($left, $rem); $donorLeft[$dk] = $left - $q; $rem -= $q
        $moves.Add([pscustomobject]@{ Откуда=$SHOPS[$d].Name; ГородОткуда=$SHOPS[$d].City; Куда=$SHOPS[$l.Store].Name; ГородКуда=$city; ОдинГород=[int]($SHOPS[$d].City -eq $city)
            Код=$l.Nom; Артикул=$nom[$l.Nom].Артикул; Номенклатура=$nom[$l.Nom].Наименование; Характеристика=$l.Har; Размер=(SizeOf $l.Har)
            ОстатокДонора=$stock[$dk]; ПродажиПолучателя13нед=$l.S13; Количество=$q })
    }
}
$moves | Export-Csv (Join-Path $PSScriptRoot 'plan_moves.csv') -NoTypeInformation -Encoding UTF8 -Delimiter "`t"
"moves: $($moves.Count) lines"

# Выход
$out = foreach ($l in $lines) {
    $st = $SHOPS[$l.Store]; $n = $nom[$l.Nom]
    [pscustomobject]@{
        СкладКод=$l.Store; Магазин=$st.Name; Город=$st.City; СрокНед=$st.Lead; Тип=$l.Type
        Группа=$n.Группа; Код=$l.Nom; Артикул=$n.Артикул; Номенклатура=$n.Наименование; Характеристика=$l.Har; Размер=(SizeOf $l.Har)
        Продано13нед=$l.S13; Продано4нед=$l.S4; СкоростьНед=$l.Rate; Остаток=$l.Stock; ВПути=$l.Transit; Цель=$l.Target; Потребность=$l.Need; КОтгрузке=$l.Qty
        ДефицитЦС=$l.Need - $l.Qty; Цена=$price[$l.Nom] + 0; СуммаОтгрузки=($price[$l.Nom] + 0) * $l.Qty; ОстатокЦС=$csStart[$l.Nom + '|' + $l.Har] + 0
    }
}
$out | Export-Csv (Join-Path $PSScriptRoot 'plan_lines.csv') -NoTypeInformation -Encoding UTF8 -Delimiter "`t"

$sum = foreach ($sc in $storeOrder) {
    $st = $SHOPS[$sc]; $my = @($out | Where-Object { $_.СкладКод -eq $sc })
    $rep = ($my | Where-Object { $_.Тип -eq 'Пополнение' } | Measure-Object КОтгрузке -Sum).Sum
    $new = ($my | Where-Object { $_.Тип -eq 'Новинка' } | Measure-Object КОтгрузке -Sum).Sum
    $def = ($my | Measure-Object ДефицитЦС -Sum).Sum
    [pscustomobject]@{ СкладКод=$sc; Магазин=$st.Name; Город=$st.City; СрокНед=$st.Lead
        Продажи13нед=$storeSales[$sc] + 0; Остаток=$storeStock[$sc] + 0
        Пополнение=$rep + 0; Новинки=$new + 0; Итого=($rep + $new); СуммаОтгрузки=($my | Measure-Object СуммаОтгрузки -Sum).Sum + 0; НетНаЦС=$def + 0; ИзДругихМагазинов=(@($moves | Where-Object { $_.Куда -eq $st.Name }) | Measure-Object Количество -Sum).Sum + 0 }
}
$sum | Export-Csv (Join-Path $PSScriptRoot 'plan_stores.csv') -NoTypeInformation -Encoding UTF8 -Delimiter "`t"
$sum | Format-Table -AutoSize | Out-String -Width 250
"lines: $($lines.Count); central start units: $(($csStart.Values | Measure-Object -Sum).Sum)"
