$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'analyze.ps1') | Out-Null
$R = [ordered]@{}
# 1. Сеть: продажи / остатки
$tSales = 0.0; foreach ($k in $sale.Keys) { $tSales += $sale[$k][0] }
$tRev = 0.0; foreach ($k in $sale.Keys) { $tRev += $sale[$k][2] }
$tStock = 0.0; foreach ($v in $stock.Values) { $tStock += $v }
$R.sales13 = $tSales; $R.rev13 = $tRev; $R.stock = $tStock; $R.weekly = $tSales / 13.57
# 2. Размеры: продаваемые размеры (продажи за 13 нед.) без остатка
$sold = 0; $zero = 0; $zeroFast = 0; $fast = 0
foreach ($k in $sale.Keys) { if ($sale[$k][0] -le 0 -or $sale[$k][2] -le 0) { continue }
    $sold++; $st = $stock[$k] + 0; $tr = $transit[$k] + 0
    if ($st + $tr -le 0) { $zero++ }
    if ($sale[$k][0] -ge 3) { $fast++; if ($st + $tr -le 0) { $zeroFast++ } } }
$R.sizesSold = $sold; $R.sizesZero = $zero; $R.sizesFast = $fast; $R.sizesFastZero = $zeroFast
# 3. Остаток без продаж: размер не продавался 13 нед.; модель-цвет не продавалась 13 нед.
$deadSize = 0.0; $deadColor = 0.0
foreach ($k in $stock.Keys) { $q = $stock[$k]
    if (-not ($sale.ContainsKey($k) -and $sale[$k][0] -gt 0)) { $deadSize += $q }
    $p = $k.Split('|', 3); if (($colorSales[$p[0] + '|' + $p[1] + '|' + (ColorOf $p[2])] + 0) -le 0) { $deadColor += $q } }
$R.deadSize = $deadSize; $R.deadColor = $deadColor
# 4. Остаток по коллекциям (Y-код в характеристике)
$byY = @{}
foreach ($k in $stock.Keys) { $h = $k.Split('|', 3)[2]; $y = if ($h -match '\bY(\d\d)\b') { 'Y' + $matches[1] } else { 'без сезона' }; $byY[$y] = $byY[$y] + $stock[$k] }
$sY = @{}
foreach ($k in $sale.Keys) { $h = $k.Split('|', 3)[2]; $y = if ($h -match '\bY(\d\d)\b') { 'Y' + $matches[1] } else { 'без сезона' }; $sY[$y] = $sY[$y] + $sale[$k][0] }
$R.byYear = @($byY.Keys | Sort-Object | % { [pscustomobject]@{ Y=$_; Stock=$byY[$_]; Sales=($sY[$_] + 0) } })
# 5. Группы
$gs = @{}; $gq = @{}
foreach ($k in $stock.Keys) { $g = GroupOf ($k.Split('|',3)[1]); $gs[$g] = $gs[$g] + $stock[$k] }
foreach ($k in $sale.Keys) { $g = GroupOf ($k.Split('|',3)[1]); $gq[$g] = $gq[$g] + $sale[$k][0] }
$R.groups = @($gs.Keys | Sort-Object { -$gs[$_] } | Select-Object -First 14 | % { [pscustomobject]@{ G=$_; Stock=$gs[$_]; Sales=($gq[$_] + 0) } })
# 6. Центральный склад по годам коллекции
$cY = @{}; foreach ($k in $csStart.Keys) { $p = $k.Split('|', 2); if (-not (IsGoods $p[0])) { continue }; $y = if ($p[1] -match '\bY(\d\d)\b') { 'Y' + $matches[1] } else { 'без сезона' }; $cY[$y] = $cY[$y] + $csStart[$k] }
$R.csByYear = @($cY.Keys | Sort-Object | % { [pscustomobject]@{ Y=$_; Q=$cY[$_] } })
$csGoods = 0.0; foreach ($k in $csStart.Keys) { if (IsGoods ($k.Split('|',2)[0])) { $csGoods += $csStart[$k] } }; $R.csGoods = $csGoods
$R | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $PSScriptRoot 'stats.json') -Encoding UTF8
$R | ConvertTo-Json -Depth 5
