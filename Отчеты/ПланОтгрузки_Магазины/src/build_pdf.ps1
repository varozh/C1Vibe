param([string]$Out)
$ErrorActionPreference = 'Stop'
$P = $PSScriptRoot
$inv = [Globalization.CultureInfo]::InvariantCulture
$ru = [Globalization.CultureInfo]::GetCultureInfo('ru-RU')
function N($s) { if (-not $s) { return 0.0 } [double]::Parse(([string]$s -replace ',', '.'), $inv) }
function F($v, $d = 0) { ([double]$v).ToString("N$d", $ru) }
$j = Get-Content "$P\stats.json" -Raw -Encoding UTF8 | ConvertFrom-Json
$stores = Import-Csv "$P\plan_stores.csv" -Delimiter "`t" -Encoding UTF8
$lines = Import-Csv "$P\plan_lines.csv" -Delimiter "`t" -Encoding UTF8
$moves = Import-Csv "$P\plan_moves.csv" -Delimiter "`t" -Encoding UTF8

$T = @{}; foreach ($r in $stores) { foreach ($c in 'Продажи13нед','Остаток','Пополнение','Новинки','Итого','СуммаОтгрузки','НетНаЦС','ИзДругихМагазинов') { $T[$c] = $T[$c] + (N $r.$c) } }
$weekly = $j.sales13 / 13.57
$coverNet = $j.stock / $weekly
$movSame = 0.0; foreach ($m in $moves) { if ($m.ОдинГород -eq '1') { $movSame += N $m.Количество } }
$def = @{}; $nm = @{}
foreach ($l in $lines) { $q = N $l.ДефицитЦС; if ($q -le 0) { continue }; $k = $l.Код + '|' + $l.Характеристика; $def[$k] = $def[$k] + $q; $nm[$k] = $l }
$mm = @{}; foreach ($m in $moves) { $k = $m.Код + '|' + $m.Характеристика; $mm[$k] = $mm[$k] + (N $m.Количество) }
$rest = 0.0; $mc = @{}
foreach ($k in $def.Keys) { $r = $def[$k] - ($mm[$k] + 0); if ($r -le 0) { continue }; $rest += $r
    $l = $nm[$k]; $c = ($l.Характеристика -replace ',\s*Размер [^,]+', '' -replace ',\s*Россия.*$', '')
    $key = $l.Номенклатура.Trim() + '|' + $c; $mc[$key] = $mc[$key] + $r }
$topDef = $mc.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10
$wave1 = @($stores | Where-Object { $_.СрокНед -eq '2' }); $wave2 = @($stores | Where-Object { $_.СрокНед -eq '1' })
function Sum($rows, $c) { $s = 0.0; foreach ($r in $rows) { $s += N $r.$c }; $s }

# --- столбчатые диаграммы (SVG) ---
function HBar($items, $max, $fmt, $color = 'var(--c1)', $w = 300) {
    $sb = New-Object Text.StringBuilder
    foreach ($it in $items) {
        $bw = [Math]::Max(1, [Math]::Round($w * $it.V / $max))
        [void]$sb.Append("<div class='hb'><span class='hl'>$($it.L)</span><svg width='$w' height='12'><rect width='$bw' height='12' rx='2' fill='$color'/></svg><span class='hv'>$(& $fmt $it.V)</span></div>")
    }
    $sb.ToString()
}
# сравнение остатка и продаж по группам: недель запаса
$grp = foreach ($g in $j.groups) { if ($g.Sales -gt 0) { [pscustomobject]@{ L = ($g.G -replace '^[\d\.]+\s*', '' -replace ' \(трикотажное изделие\)', ''); Stock = $g.Stock; W = $g.Stock / ($g.Sales / 13.57) } } }
$grpRows = ($grp | ForEach-Object { "<tr><td>$($_.L)</td><td class='n'>$(F $_.Stock)</td><td class='n'>$(F ($_.Stock / $_.W))</td><td class='n'><b>$(F $_.W)</b></td></tr>" }) -join ''
# коллекции
$yearMap = @{}; foreach ($y in $j.byYear) { $yearMap[$y.Y] = $y }
$old = 0.0; $oldS = 0.0; foreach ($y in $j.byYear) { if ($y.Y -match '^Y(\d\d)$' -and [int]$matches[1] -le 24) { $old += $y.Stock; $oldS += $y.Sales } }
$coll = @(
  [pscustomobject]@{ L='Y21–Y24 (прошлые сезоны)'; S=$old; Q=$oldS },
  [pscustomobject]@{ L='Y25'; S=$yearMap['Y25'].Stock; Q=$yearMap['Y25'].Sales },
  [pscustomobject]@{ L='Y26 (текущий)'; S=$yearMap['Y26'].Stock; Q=$yearMap['Y26'].Sales },
  [pscustomobject]@{ L='Y27 (новая зима)'; S=$yearMap['Y27'].Stock; Q=$yearMap['Y27'].Sales },
  [pscustomobject]@{ L='Аксессуары, бижутерия (без сезона)'; S=$yearMap['без сезона'].Stock; Q=$yearMap['без сезона'].Sales }
)
$collRows = ($coll | ForEach-Object { "<tr><td>$($_.L)</td><td class='n'>$(F $_.S)</td><td class='n'>$(F $_.Q)</td><td class='n'><b>$(F ($_.S / ($_.Q / 13.57)))</b></td></tr>" }) -join ''
$cs = @{}; foreach ($y in $j.csByYear) { $cs[$y.Y] = $y.Q }
$csOld = 0.0; foreach ($y in $j.csByYear) { if ($y.Y -match '^Y(\d\d)$' -and [int]$matches[1] -le 25) { $csOld += $y.Q } }

# таблица магазинов
function StoreRows($rows) {
    ($rows | Sort-Object @{e={-(N $_.Продажи13нед)}} | ForEach-Object {
        $wk = (N $_.Продажи13нед) / 13.57; $cov = if ($wk -gt 0) { (N $_.Остаток) / $wk } else { 0 }
        "<tr><td>$($_.Магазин)</td><td>$($_.Город)</td><td class='n'>$(F $wk)</td><td class='n'>$(F $_.Остаток)</td><td class='n'>$(F $cov)</td><td class='n'>$(F $_.Пополнение)</td><td class='n'>$(F $_.Новинки)</td><td class='n'><b>$(F $_.Итого)</b></td><td class='n'>$(F ((N $_.СуммаОтгрузки)/1e6) 2)</td><td class='n'>$(F $_.ИзДругихМагазинов)</td></tr>"
    }) -join ''
}
$storeBars = HBar ($stores | Sort-Object @{e={-(N $_.Итого)}} | ForEach-Object { [pscustomobject]@{ L=$_.Магазин; V=(N $_.Итого) } }) ([double](($stores | ForEach-Object { N $_.Итого } | Measure-Object -Maximum).Maximum)) { param($v) F $v } 'var(--c1)' 260
$defRows = ($topDef | ForEach-Object { $pp = $_.Key.Split('|', 2); "<tr><td>$($pp[0])</td><td>$($pp[1])</td><td class='n'><b>$(F $_.Value)</b></td></tr>" }) -join ''

$sizePct = 100.0 * $j.sizesZero / $j.sizesSold
$deadSizePct = 100.0 * $j.deadSize / $j.stock
$deadColorPct = 100.0 * $j.deadColor / $j.stock

$html = @"
<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>Анализ остатков и план отгрузки</title>
<style>
:root { --ink:#1f2328; --mut:#5b6470; --line:#d9dee4; --bg2:#f4f6f8; --c1:#2f6db5; --c2:#c8553d; --c3:#3a8a5c; }
@page { size: A4; margin: 14mm 13mm 14mm 13mm; }
* { box-sizing: border-box; }
body { font-family: 'Segoe UI', Arial, sans-serif; color: var(--ink); font-size: 10pt; line-height: 1.42; margin: 0; background:#fff; }
h1 { font-size: 19pt; margin: 0 0 2mm; letter-spacing: -0.2px; }
h2 { font-size: 13pt; margin: 7mm 0 2.5mm; padding-bottom: 1.2mm; border-bottom: 1.5px solid var(--ink); }
h3 { font-size: 10.5pt; margin: 4mm 0 1.5mm; }
p { margin: 0 0 2mm; } h2, h3 { break-after: avoid; page-break-after: avoid; }
.sub { color: var(--mut); margin-bottom: 5mm; }
.kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 3mm; margin: 3mm 0 4mm; }
.kpi { background: var(--bg2); border-radius: 2mm; padding: 3mm 3.5mm; }
.kpi .v { font-size: 17pt; font-weight: 700; line-height: 1.1; }
.kpi .l { color: var(--mut); font-size: 8.5pt; margin-top: 1mm; }
.kpi.alert .v { color: var(--c2); }
.kpi.ok .v { color: var(--c3); }
.box { border-left: 3px solid var(--c1); background: #f2f6fb; padding: 2.5mm 4mm; margin: 3mm 0; }
.box.warn { border-color: var(--c2); background: #fbf3f1; }
table { width: 100%; border-collapse: collapse; font-size: 8.6pt; margin: 1.5mm 0 3mm; }
th { text-align: left; font-weight: 600; color: var(--mut); border-bottom: 1px solid var(--ink); padding: 1.2mm 1.5mm; vertical-align: bottom; }
td { border-bottom: 1px solid var(--line); padding: 1mm 1.5mm; }
table.st td:first-child { white-space: nowrap; } table.st td:nth-child(2) { white-space: nowrap; }
td.n, th.n { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
tr.tot td { font-weight: 700; border-top: 1px solid var(--ink); border-bottom: none; }
ul { margin: 1mm 0 2mm 5mm; padding: 0; } li { margin-bottom: 1mm; }
.two { display: grid; grid-template-columns: 1fr 1fr; gap: 6mm; }
.hb { display: flex; align-items: center; gap: 2mm; font-size: 8.2pt; margin: 0.6mm 0; }
.hb .hl { width: 46mm; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.hb .hv { font-variant-numeric: tabular-nums; color: var(--mut); }
.pb { break-before: page; }
.avoid { break-inside: avoid; }
.foot { color: var(--mut); font-size: 8pt; margin-top: 6mm; }
.wave { display:grid; grid-template-columns: 1fr 1fr; gap: 4mm; margin: 2mm 0 3mm; }
.wave div { border: 1px solid var(--line); border-radius: 2mm; padding: 3mm 3.5mm; }
.wave b.t { display:block; font-size: 11pt; margin-bottom: 1mm; }
</style></head><body>

<h1>Остатки, продажи и план отгрузки в магазины Business Line</h1>
<div class="sub">Данные 1С:Розница и 1С:УТ на 19.11.2026 · продажи за 13,6 недели (17.08–19.11.2026) · 33 магазина Business Line и «ОФИС» (E-COM)</div>

<h2>Главное</h2>
<div class="kpis">
  <div class="kpi alert"><div class="v">$(F $coverNet) нед.</div><div class="l">запас товара в магазинах при текущих продажах ($(F $j.stock) шт. на полках при $(F $weekly) шт. продаж в неделю)</div></div>
  <div class="kpi alert"><div class="v">$(F $sizePct) %</div><div class="l">размеров, проданных магазином за 13 недель, сейчас в этом магазине нет ($(F $j.sizesZero) из $(F $j.sizesSold))</div></div>
  <div class="kpi"><div class="v">$(F $T.Итого) шт.</div><div class="l">предлагается отгрузить с центрального склада, около $(F ($T.СуммаОтгрузки/1e6) 1) млн ₽ по розничной цене</div></div>
  <div class="kpi"><div class="v">$(F $rest) шт.</div><div class="l">не хватает ни на складе, ни в других магазинах: нужен дошив или закупка</div></div>
</div>
<ul>
  <li><b>Товара в магазинах много, но не того.</b> Магазин держит 1,9–3 тыс. шт. при 25–100 продажах в неделю, то есть запаса хватит на 26–100 недель. При этом $(F $deadSizePct) % остатка — размеры, которые не продавались в этом магазине 13 недель, а $(F $deadColorPct) % — модели и цвета без единой продажи в этом магазине за 13 недель.</li>
  <li><b>Хорошо продающиеся размеры заканчиваются.</b> Обычно магазин получает по 1 шт. размера: размер продаётся — и его больше нет. Поэтому главная задача отгрузки — не объём, а возврат продаваемых размеров на полку.</li>
  <li><b>Центральный склад не закрывает весь спрос.</b> Из $(F ($T.Пополнение + $T.НетНаЦС)) шт. потребности на пополнение склад покрывает $(F $T.Пополнение) шт. Ещё $(F $T.ИзДругихМагазинов) шт. можно взять в магазинах, где эти размеры лежат без продаж ($(F $movSame) шт. — внутри того же города), а $(F $rest) шт. нужно дошить.</li>
  <li><b>Перед Новым годом продажи растут.</b> В декабре 2025 они были на 18–19 % выше, чем в конце октября — ноябре, поэтому запас на декабрь в расчёте увеличен в 1,2 раза.</li>
</ul>

<h2>1. Товара в магазинах много</h2>
<p>На полках 33 магазинов и «ОФИСа» $(F $j.stock) шт. товара (без упаковки, вешалок, оборудования и сертификатов). За 13,6 недели продано $(F $j.sales13) шт. на $(F ($j.rev13/1e6) 1) млн ₽, то есть $(F $weekly) шт. в неделю. Запаса хватит на <b>$(F $coverNet) нед.</b>, почти на год. Для одежды нормальный запас в магазине — 8–12 недель.</p>
<div class="two avoid">
<div><h3>Запас по группам товаров</h3>
<table><tr><th>Группа</th><th class="n">Остаток, шт</th><th class="n">Продажи в нед.</th><th class="n">Запас, нед.</th></tr>$grpRows</table></div>
<div><h3>Запас по коллекциям</h3>
<table><tr><th>Коллекция</th><th class="n">Остаток, шт</th><th class="n">Продано 13 нед.</th><th class="n">Запас, нед.</th></tr>$collRows</table>
<div class="box warn">$(F $old) шт. коллекций Y21–Y24 и $(F $oldS) шт. продаж за 13 недель: их запаса хватит на $(F ($old/($oldS/13.57))) нед. Такой товар не продастся без уценки, и его стоит перевести в дисконт или вернуть на склад.</div>
<p>На центральном складе $(F $j.csGoods) шт. товара, из них $(F $csOld) шт. — коллекции Y25 и старше, $(F $cs['Y26']) шт. — Y26 и всего $(F $cs['Y27']) шт. новой Y27.</p>
</div></div>

<h2>2. Хорошо продающиеся размеры заканчиваются</h2>
<p>Каждый размер модели магазин обычно получает в 1–2 экземплярах. Из $(F $j.sizesSold) размеров, которые магазины продали за 13 недель, <b>$(F $j.sizesZero) ($(F $sizePct) %) сейчас в этом магазине нет</b>, и в пути их тоже нет. Тем временем $(F $j.deadSize) шт. ($(F $deadSizePct) % остатка) — размеры, которые в этом магазине за 13 недель не продались ни разу.</p>
<div class="box">Вывод: в первую очередь надо возвращать на полку проданные размеры, а не наращивать общий объём. Залежавшиеся размеры стоит передать туда, где они продаются.</div>

<h2>3. Как рассчитан план</h2>
<ul>
  <li><b>Скорость продаж</b> считается по каждому размеру модели в каждом магазине: наполовину по последним 4 неделям, наполовину по 13 неделям. Для новых магазинов (Волгоград, Члб Родник) — за фактический срок работы.</li>
  <li><b>Цель по размеру</b> = скорость × 1,2 (декабрь) × (срок доставки + 1 неделя до следующей отгрузки), но не меньше 1 шт., если размер продавался. <b>К отгрузке</b> = цель − остаток − товар в пути.</li>
  <li><b>Срок доставки:</b> 2 недели — Екатеринбург, Пермь, Челябинск, Сочи, Санкт-Петербург; 1 неделя — остальные, включая Уфу и Оренбург.</li>
  <li><b>Если товара на складе мало,</b> его сначала получают магазины, где размера нет совсем, затем — где он продаётся быстрее.</li>
  <li><b>Новинки:</b> модели, проданные по сети не меньше 8 шт. за 4 недели, которых в магазине нет: размерный ряд 42–54 по 1 шт., до 12 моделей на магазин.</li>
</ul>

<h2>4. План отгрузки</h2>
<div class="wave">
  <div><b class="t">Волна 1 — отгрузить 20–21.11</b>Екатеринбург (3), Пермь, Челябинск (2), Сочи, Санкт-Петербург (2): $(F $wave1.Count) магазинов, $(F (Sum $wave1 'Итого')) шт., $(F ((Sum $wave1 'СуммаОтгрузки')/1e6) 1) млн ₽. Поступление 01–05.12.</div>
  <div><b class="t">Волна 2 — отгрузить 24–25.11</b>Остальные $(F $wave2.Count) магазинов, включая Уфу и Оренбург: $(F (Sum $wave2 'Итого')) шт., $(F ((Sum $wave2 'СуммаОтгрузки')/1e6) 1) млн ₽. Поступление 01–02.12.</div>
</div>
<p>Если нужна ещё одна поставка к новогоднему пику (20–31.12), отправлять её в дальние магазины лучше до 08.12, в остальные — до 15.12.</p>
<h3>Волна 1 — доставка 1,5–2 недели</h3>
<table class='st'><tr><th>Магазин</th><th>Город</th><th class="n">Продажи в нед.</th><th class="n">Остаток</th><th class="n">Запас, нед.</th><th class="n">Пополнение</th><th class="n">Новинки</th><th class="n">Итого, шт</th><th class="n">Сумма, млн ₽</th><th class="n">Из др. магазинов</th></tr>
$(StoreRows $wave1)
<tr class="tot"><td>Итого волна 1</td><td></td><td class="n">$(F ((Sum $wave1 'Продажи13нед')/13.57))</td><td class="n">$(F (Sum $wave1 'Остаток'))</td><td></td><td class="n">$(F (Sum $wave1 'Пополнение'))</td><td class="n">$(F (Sum $wave1 'Новинки'))</td><td class="n">$(F (Sum $wave1 'Итого'))</td><td class="n">$(F ((Sum $wave1 'СуммаОтгрузки')/1e6) 2)</td><td class="n">$(F (Sum $wave1 'ИзДругихМагазинов'))</td></tr></table>
<h3>Волна 2 — доставка 1 неделя</h3>
<table class='st'><tr><th>Магазин</th><th>Город</th><th class="n">Продажи в нед.</th><th class="n">Остаток</th><th class="n">Запас, нед.</th><th class="n">Пополнение</th><th class="n">Новинки</th><th class="n">Итого, шт</th><th class="n">Сумма, млн ₽</th><th class="n">Из др. магазинов</th></tr>
$(StoreRows $wave2)
<tr class="tot"><td>Итого волна 2</td><td></td><td class="n">$(F ((Sum $wave2 'Продажи13нед')/13.57))</td><td class="n">$(F (Sum $wave2 'Остаток'))</td><td></td><td class="n">$(F (Sum $wave2 'Пополнение'))</td><td class="n">$(F (Sum $wave2 'Новинки'))</td><td class="n">$(F (Sum $wave2 'Итого'))</td><td class="n">$(F ((Sum $wave2 'СуммаОтгрузки')/1e6) 2)</td><td class="n">$(F (Sum $wave2 'ИзДругихМагазинов'))</td></tr></table>

<div class="avoid"><h3>Отгрузка с центрального склада по магазинам, шт.</h3>$storeBars</div>

<h2>5. Чего не хватает на центральном складе</h2>
<p>Из потребности на пополнение $(F ($T.Пополнение + $T.НетНаЦС)) шт. на складе нет $(F $T.НетНаЦС) шт.:</p>
<ul>
  <li><b>$(F $T.ИзДругихМагазинов) шт. — перемещения между магазинами.</b> Берётся размер, который у магазина-донора лежит 13 недель без продаж (и модель там продаётся не больше 1 шт.). Из них $(F $movSame) шт. — внутри одного города, их можно передать быстро. Остальные — через центральный склад или прямым перемещением.</li>
  <li><b>$(F $rest) шт. — дошив или закупка.</b> Больше всего не хватает моделей текущей коллекции Y26:</li>
</ul>
<table><tr><th>Модель</th><th>Цвет / ткань</th><th class="n">Не хватает, шт</th></tr>$defRows</table>

<h2>6. Рекомендации</h2>
<ul>
  <li>Отгрузить волну 1 до 21.11, волну 2 — до 25.11 (лист «План отгрузки» в Excel).</li>
  <li>Оформить перемещения внутри городов ($(F $movSame) шт.): Москва, Екатеринбург, Саратов, Самара, Казань, Нижний Новгород, Пенза и другие (лист «Перемещения м-ду магазинами»).</li>
  <li>Передать в производство список размеров для дошива (лист «Дефицит ЦС»): прежде всего жакеты, брюки, блузы и юбки Y26 размеров 44–50.</li>
  <li>Остатки прошлых коллекций Y21–Y24 ($(F $old) шт. в магазинах) перевести в дисконт или вернуть на склад и не отгружать их в магазины повторно.</li>
  <li>Перед Новым годом повторить расчёт через неделю на свежих данных: дальние магазины — до 08.12, ближние — до 15.12.</li>
</ul>
<div class="box warn">Ограничения: резервы под интернет-магазин и маркетплейсы на центральном складе не вычитались. У части магазинов-доноров могут быть недавно поступившие вещи, которые ещё не успели продаться, поэтому список перемещений стоит просмотреть перед оформлением.</div>
<div class="foot">Построчный план: «План_отгрузки_2026-11-19.xlsx» (листы «План отгрузки», «Перемещения м-ду магазинами», «Дефицит ЦС»). Расчёт: скрипты в папке src.</div>
</body></html>
"@
$hp = Join-Path $P 'report.html'
[IO.File]::WriteAllText($hp, $html, (New-Object Text.UTF8Encoding($false)))
$edge = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
if (Test-Path $Out) { Remove-Item $Out -Force }
$uri = ([Uri]$hp).AbsoluteUri
$pr = Start-Process $edge -ArgumentList "--headless=new --disable-gpu --no-pdf-header-footer --user-data-dir=`"$P\edge-prof`" --print-to-pdf=`"$Out`" `"$uri`"" -PassThru -Wait -WindowStyle Hidden
"pdf: $(Test-Path $Out) $((Get-Item $Out -ErrorAction SilentlyContinue).Length)"
