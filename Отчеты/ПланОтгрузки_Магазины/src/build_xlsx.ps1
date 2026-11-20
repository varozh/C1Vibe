param([string]$Out)
$ErrorActionPreference = 'Stop'
$P = $PSScriptRoot
$inv = [Globalization.CultureInfo]::InvariantCulture
function N($s) { if (-not $s) { return 0.0 } [double]::Parse(($s -replace ',', '.'), $inv) }
$lines = Import-Csv "$P\plan_lines.csv" -Delimiter "`t" -Encoding UTF8
$stores = Import-Csv "$P\plan_stores.csv" -Delimiter "`t" -Encoding UTF8
$moves = @(Import-Csv "$P\plan_moves.csv" -Delimiter "`t" -Encoding UTF8)

$wave = @{ '2' = 'Волна 1: 20–21.11'; '1' = 'Волна 2: 24–25.11' }
$arr  = @{ '2' = '01–05.12'; '1' = '01–02.12' }

function Put($ws, $header, $data, $widths) {
    $r = $data.Count; $c = $header.Count
    $a = New-Object 'object[,]' ($r + 1), $c
    for ($j = 0; $j -lt $c; $j++) { $a[0, $j] = $header[$j] }
    for ($i = 0; $i -lt $r; $i++) { for ($j = 0; $j -lt $c; $j++) { $a[($i + 1), $j] = $data[$i][$j] } }
    $rng = $ws.Range($ws.Cells(1, 1), $ws.Cells($r + 1, $c)); $rng.Value2 = $a
    $h = $ws.Range($ws.Cells(1, 1), $ws.Cells(1, $c)); $h.Font.Bold = $true; $h.Interior.Color = 0xEFE3D9; $h.WrapText = $true
    $ws.Rows(1).RowHeight = 32
    for ($j = 0; $j -lt $c; $j++) { $ws.Columns($j + 1).ColumnWidth = $widths[$j] }
    [void]$rng.AutoFilter()
    $ws.Activate(); $ws.Application.ActiveWindow.SplitRow = 1; $ws.Application.ActiveWindow.FreezePanes = $true
}

if (-not ('OleMsgFilter' -as [type])) {
Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IOleMessageFilter { [PreserveSig] int HandleInComingCall(int a, IntPtr b, int c, IntPtr d); [PreserveSig] int RetryRejectedCall(IntPtr a, int b, int c); [PreserveSig] int MessagePending(IntPtr a, int b, int c); }
public class OleMsgFilter : IOleMessageFilter {
  [DllImport("Ole32.dll")] static extern int CoRegisterMessageFilter(IOleMessageFilter n, out IOleMessageFilter o);
  public static void Register() { IOleMessageFilter o; CoRegisterMessageFilter(new OleMsgFilter(), out o); }
  public static void Revoke() { IOleMessageFilter o; CoRegisterMessageFilter(null, out o); }
  int IOleMessageFilter.HandleInComingCall(int a, IntPtr b, int c, IntPtr d) { return 0; }
  int IOleMessageFilter.RetryRejectedCall(IntPtr a, int b, int c) { if (c == 2 && b < 600000) return 200; return -1; }
  int IOleMessageFilter.MessagePending(IntPtr a, int b, int c) { return 2; }
}
"@
}
[OleMsgFilter]::Register()
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
while ($wb.Worksheets.Count -lt 5) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }

# 1. Итоги
$ws = $wb.Worksheets.Item(1); $ws.Name = 'Итоги по магазинам'
$data = foreach ($s in ($stores | Sort-Object @{e={-(N $_.СрокНед)}}, @{e={-(N $_.Продажи13нед)}})) {
    $wk = (N $s.Продажи13нед) / 13.57
    $cover = if ($wk -gt 0) { [Math]::Round((N $s.Остаток) / $wk, 1) } else { $null }
    $lead = if ($s.СрокНед -eq '2') { '1,5–2 нед.' } else { '1 нед.' }
    ,@($wave[$s.СрокНед], $s.Магазин, $s.Город, $lead, $arr[$s.СрокНед], (N $s.Продажи13нед), [Math]::Round($wk, 0), (N $s.Остаток), $cover,
       (N $s.Пополнение), (N $s.Новинки), (N $s.Итого), (N $s.СуммаОтгрузки), (N $s.ИзДругихМагазинов), (N $s.НетНаЦС))
}
$data = @($data)
Put $ws @('Волна отгрузки','Магазин','Город','Доставка','Ожидаемое поступление','Продано за 13 нед., шт','Продаж в неделю, шт','Остаток сейчас, шт','Запас, недель','Пополнение с ЦС, шт','Новинки с ЦС, шт','Итого отгрузка с ЦС, шт','Сумма отгрузки (по рознич. цене), ₽','Из других магазинов, шт','Не хватает на ЦС, шт') $data @(18,28,16,10,12,11,10,11,9,11,10,11,14,11,11)
$n = $data.Count + 1
$ws.Cells($n + 1, 2).Value2 = 'ИТОГО'; $ws.Cells($n + 1, 2).Font.Bold = $true
foreach ($col in 6,7,8,10,11,12,13,14,15) { $c = [string][char](64 + $col); $ws.Cells($n + 1, $col).Formula = "=SUBTOTAL(9,${c}2:${c}$n)" }
$ws.Range("M2:M$($n + 1)").NumberFormat = '# ##0'
$ws.Range("F2:O$($n + 1)").HorizontalAlignment = -4152
$ws.Rows($n + 1).Font.Bold = $true

# 2. План отгрузки
$ws = $wb.Worksheets.Item(2); $ws.Name = 'План отгрузки'
$sel = $lines | Where-Object { (N $_.КОтгрузке) -gt 0 } | Sort-Object @{e={-(N $_.СрокНед)}}, Магазин, Тип, Группа, Артикул, Характеристика
$data = @(foreach ($l in $sel) {
    ,@($wave[$l.СрокНед], $l.Магазин, $l.Город, $l.Тип, $l.Группа, $l.Код, $l.Артикул, $l.Номенклатура, $l.Характеристика, $l.Размер,
       (N $l.Продано13нед), (N $l.Продано4нед), (N $l.СкоростьНед), (N $l.Остаток), (N $l.ВПути), (N $l.Цель), (N $l.КОтгрузке), (N $l.Цена), (N $l.СуммаОтгрузки))
})
Put $ws @('Волна','Магазин','Город','Тип','Группа','Код','Артикул','Номенклатура','Характеристика','Размер','Продано 13 нед.','Продано 4 нед.','Скорость, шт/нед','Остаток','В пути','Цель','К отгрузке','Цена','Сумма') $data @(17,24,14,11,20,12,10,30,50,8,8,8,8,8,7,6,8,9,10)

# 3. Перемещения
$ws = $wb.Worksheets.Item(3); $ws.Name = 'Перемещения м-ду магазинами'
$msel = $moves | Sort-Object @{e={-(N $_.ОдинГород)}}, ГородКуда, Куда, Артикул, Характеристика
$data = @(foreach ($m in $msel) {
    ,@($(if ($m.ОдинГород -eq '1') { 'Да' } else { 'Нет' }), $m.Откуда, $m.ГородОткуда, $m.Куда, $m.ГородКуда, $m.Код, $m.Артикул, $m.Номенклатура, $m.Характеристика, $m.Размер, (N $m.ОстатокДонора), (N $m.ПродажиПолучателя13нед), (N $m.Количество))
})
Put $ws @('В пределах города','Откуда (нет продаж 13 нед.)','Город откуда','Куда','Город куда','Код','Артикул','Номенклатура','Характеристика','Размер','Остаток у донора','Продажи получателя 13 нед.','Количество') $data @(10,24,14,24,14,12,10,30,50,8,9,10,9)

# 4. Дефицит ЦС (после перемещений)
$ws = $wb.Worksheets.Item(4); $ws.Name = 'Дефицит ЦС'
$def = @{}; $info = @{}
foreach ($l in $lines) { $d = N $l.ДефицитЦС; if ($d -le 0) { continue }
    $k = $l.Код + '|' + $l.Характеристика; $def[$k] = $def[$k] + $d
    if (-not $info[$k]) { $info[$k] = @($l.Группа, $l.Код, $l.Артикул, $l.Номенклатура, $l.Характеристика, $l.Размер, 0, 0.0, (N $l.Цена)) }
    $info[$k][6] = $info[$k][6] + 1; $info[$k][7] = $info[$k][7] + (N $l.Продано13нед) }
$mv = @{}; foreach ($m in $moves) { $k = $m.Код + '|' + $m.Характеристика; $mv[$k] = $mv[$k] + (N $m.Количество) }
$data = @(foreach ($k in ($def.Keys | Sort-Object @{e={-($def[$_] - $mv[$_])}})) {
    $i = $info[$k]; $rest = $def[$k] - ($mv[$k] + 0)
    ,@($i[0], $i[1], $i[2], $i[3], $i[4], $i[5], $i[6], $i[7], $def[$k], ($mv[$k] + 0), $rest, $i[8])
})
Put $ws @('Группа','Код','Артикул','Номенклатура','Характеристика','Размер','Магазинов с потребностью','Продано в этих магазинах 13 нед.','Потребность сверх ЦС','Закрывается перемещениями','Остаётся (дошив / закупка)','Цена') $data @(20,12,10,30,50,8,11,12,11,12,12,9)

# 5. Методика
$ws = $wb.Worksheets.Item(5); $ws.Name = 'Методика'
$txt = @(
 'План отгрузки товаров в магазины Business Line — расчёт на 19.11.2026',
 '',
 'Данные: 1С:Розница (bl-isoflex/retail) — продажи 17.08–19.11.2026 (13,6 нед.), остатки и товары в пути магазинов на 19.11; 1С:УТ (bl-isoflex/trade) — остаток «Главного розничного склада» (качество «Новый»).',
 'Магазины: 33 действующих магазина Business Line и «ОФИС» (E-COM). Закрытые, SoLi''s, Gesoline и помеченные на удаление не учитываются. Сертификаты, оборудование, упаковка и подарочные позиции с нулевой ценой не входят в план.',
 'Скорость продаж: 50% — средняя за последние 4 недели, 50% — за 13 недель (для новых магазинов — за фактический срок работы), в разрезе модель / цвет / размер.',
 'Сезонность: ×1,2 — в декабре 2025 продажи были на 18–19% выше, чем в конце октября — ноябре.',
 'Цель по каждому размеру = скорость × 1,2 × (срок доставки + 1 неделя до следующей отгрузки), но не меньше 1 шт., если размер продавался за 13 недель. К отгрузке = цель − остаток − в пути.',
 'Срок доставки: 1 неделя; 2 недели — Екатеринбург, Пермь, Челябинск (Урал), Сочи, Санкт-Петербург.',
 'Если на складе не хватает товара, в первую очередь его получают магазины, где размера нет совсем; затем — по скорости продаж; при равенстве — дальние магазины.',
 'Новинки: модели и цвета, проданные по сети не меньше 8 шт. за 4 недели, которых нет в магазине, — по 1 шт. размеров 42–54 из наличия ЦС, до 12 моделей на магазин, начиная с магазинов с большими продажами.',
 'Перемещения: размер, которого не хватает на ЦС, берётся из магазина, где он лежит 13 недель без продаж (и модель-цвет там продаётся не больше 1 шт.); сначала — из того же города.',
 'Волна 1 (дальние магазины) — отгрузить 20–21.11, поступление 01–05.12. Волна 2 — отгрузить 24–25.11, поступление 01–02.12. Для новогоднего пика последняя отгрузка в дальние магазины — до 08.12, в остальные — до 15.12.'
)
for ($i = 0; $i -lt $txt.Count; $i++) { $ws.Cells($i + 1, 1).Value2 = $txt[$i] }
$ws.Columns(1).ColumnWidth = 160; $ws.Range("A1:A$($txt.Count)").WrapText = $true; $ws.Cells(1, 1).Font.Bold = $true; $ws.Cells(1, 1).Font.Size = 13

$wb.Worksheets.Item(1).Activate()
if (Test-Path $Out) { Remove-Item $Out -Force }
$wb.SaveAs($Out, 51)
$wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
"saved $Out; plan rows $($sel.Count); moves $($moves.Count); deficit SKUs $($def.Count)"
