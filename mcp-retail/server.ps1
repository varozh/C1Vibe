# MCP-сервер (stdio) для чтения данных 1С через COM-соединение.
# Разрядность PowerShell должна совпадать с зарегистрированным COM-коннектором нужной версии платформы:
# Розница (8.3.12, 32 бита) — C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe,
# ЗУП (8.3.27, 64 бита) — C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe.
# Структура метаданных берётся из XML-выгрузки конфигурации (если в конфиге задан dump) или из метаданных базы через COM,
# данные — запросами через V83.COMConnector.
param([string]$Config = (Join-Path $PSScriptRoot 'config.json'))

$ErrorActionPreference = 'Stop'
$utf8 = New-Object Text.UTF8Encoding($false)
$stdin = New-Object IO.StreamReader([Console]::OpenStandardInput(), $utf8)
$stdout = New-Object IO.StreamWriter([Console]::OpenStandardOutput(), $utf8)
$stdout.AutoFlush = $true
$inv = [Globalization.CultureInfo]::InvariantCulture

$cfg = Get-Content $Config -Raw -Encoding UTF8 | ConvertFrom-Json
$serverName = if ($cfg.name) { [string]$cfg.name } else { 'retail-1c' }
$instructions = if ($cfg.instructions) { [string]$cfg.instructions } else { 'Чтение данных 1С:Розница 1.0. Сначала describe (find/object), чтобы узнать точные имена таблиц и полей, затем query.' }
$queryNote = if ($null -ne $cfg.queryNote) { [string]$cfg.queryNote } else { 'Конфигурация в режиме совместимости 8.1 — используй синтаксис запросов 8.1/8.2.' }

function Write-Log([string]$msg) { [Console]::Error.WriteLine("[$serverName] $msg") }

#region Подключение к базе

$script:db = $null

# Адаптер COM в PowerShell не видит свойства объектов 1С, поэтому все обращения идут через позднее связывание.
# Запятая перед результатом не даёт конвейеру развернуть коллекции 1С.
$comType = [System.__ComObject]
function Get-ComProp($o, [string]$name) { , $comType.InvokeMember($name, 'GetProperty', $null, $o, $null) }
function Set-ComProp($o, [string]$name, $value) { [void]$comType.InvokeMember($name, 'SetProperty', $null, $o, @(, $value)) }
function Invoke-ComMethod($o, [string]$name) { , $comType.InvokeMember($name, 'InvokeMethod', $null, $o, [object[]]$args) }

# Пароль в конфиге хранится зашифрованным (passwordEnc) через Windows DPAPI: расшифровать его может только
# тот же пользователь Windows на том же компьютере. Зашифровать: protect-password.ps1 -Config <путь к config.json>.
function Get-Password {
    if (-not $cfg.passwordEnc) { return [string]$cfg.password }
    try { $ss = ConvertTo-SecureString ([string]$cfg.passwordEnc) }
    catch { throw "Не удалось расшифровать passwordEnc: он зашифрован другим пользователем Windows или на другом компьютере. Зашифруйте пароль заново: protect-password.ps1 -Config `"$Config`"" }
    $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
}

function Connect-Db {
    if ($null -ne $script:db) { return }
    $q = { param($s) ([string]$s).Replace('"', '""') }
    if ($cfg.file) { $conn = 'File="{0}";' -f (& $q $cfg.file) }
    else { $conn = 'Srvr="{0}";Ref="{1}";' -f (& $q $cfg.server), (& $q $cfg.ref) }
    if ($cfg.user) { $conn += 'Usr="{0}";Pwd="{1}";' -f (& $q $cfg.user), (& $q (Get-Password)) }
    try { $connector = New-Object -ComObject V83.COMConnector }
    catch { throw "COM-коннектор 1С не зарегистрирован для $([IntPtr]::Size * 8)-битного процесса. Выполните от администратора regsvr32 для comcntr.dll нужной версии платформы (32 бита: C:\Windows\SysWOW64\regsvr32.exe, 64 бита: C:\Windows\System32\regsvr32.exe)." }
    $script:db = $connector.Connect($conn)
    Write-Log "Подключено: $($conn -replace 'Pwd="[^"]*"', 'Pwd="***"')"
}

function Format-Value($v) {
    if ($null -eq $v -or $v -is [DBNull]) { return '' }
    # Пустая дата 1С приходит через COM как 0001-01-01 (8.3.12) или 0100-01-01 (8.3.27)
    if ($v -is [datetime]) { if ($v.Year -le 100) { return '' } return $v.ToString('yyyy-MM-dd HH:mm:ss') }
    if ($v -is [bool]) { if ($v) { return 'true' } return 'false' }
    if ($v -is [string]) { return ($v -replace '[\t\r\n]+', ' ') }
    if ($v -is [ValueType]) { return [Convert]::ToString($v, $inv) }
    return ([string](Invoke-ComMethod $script:db 'String' $v)) -replace '[\t\r\n]+', ' '
}

function ConvertTo-1CParam($v) {
    if ($v -is [string] -and $v -match '^\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2})?)?$') {
        return [datetime]::Parse($v, $inv)
    }
    if ($v -is [array]) {
        $arr = Invoke-ComMethod $script:db 'NewObject' 'Array'
        foreach ($item in $v) { [void](Invoke-ComMethod $arr 'Add' (ConvertTo-1CParam $item)) }
        return , $arr
    }
    return $v
}

function Invoke-1CQuery($a) {
    Connect-Db
    $limit = 100
    if ($a.limit) { $limit = [Math]::Min([int]$a.limit, 5000) }
    $q = Invoke-ComMethod $script:db 'NewObject' 'Query'
    Set-ComProp $q 'Text' ([string]$a.text)
    if ($a.params) {
        foreach ($p in $a.params.PSObject.Properties) { [void](Invoke-ComMethod $q 'SetParameter' $p.Name (ConvertTo-1CParam $p.Value)) }
    }
    $res = Invoke-ComMethod $q 'Execute'
    $columns = Get-ComProp $res 'Columns'
    $n = Invoke-ComMethod $columns 'Count'
    $cols = for ($i = 0; $i -lt $n; $i++) { Get-ComProp (Invoke-ComMethod $columns 'Get' $i) 'Name' }
    $sel = Invoke-ComMethod $res 'Select'
    $total = Invoke-ComMethod $sel 'Count'
    $sb = New-Object Text.StringBuilder
    [void]$sb.AppendLine((@($cols) -join "`t"))
    $shown = 0
    while ($shown -lt $limit -and (Invoke-ComMethod $sel 'Next')) {
        $vals = for ($i = 0; $i -lt $n; $i++) { Format-Value (Invoke-ComMethod $sel 'Get' $i) }
        [void]$sb.AppendLine((@($vals) -join "`t"))
        $shown++
    }
    [void]$sb.Append("-- строк: $shown из $total")
    if ($shown -lt $total) { [void]$sb.Append(" (обрезано по limit=$limit)") }
    return $sb.ToString()
}

function Get-DbInfo {
    Connect-Db
    $md = Get-ComProp $script:db 'Metadata'
    $target = if ($cfg.file) { "File=$($cfg.file)" } else { "$($cfg.server)/$($cfg.ref)" }
    return "База: $target`nКонфигурация: $(Get-ComProp $md 'Name') ($(Get-ComProp $md 'Synonym')), версия $(Get-ComProp $md 'Version')`nПользователь: $(Invoke-ComMethod $script:db 'UserName')"
}

#endregion

#region Метаданные

# Вид объекта -> папка выгрузки (она же имя коллекции в Метаданные через COM)
$kinds = [ordered]@{
    'Справочник'             = 'Catalogs'
    'Документ'               = 'Documents'
    'РегистрНакопления'      = 'AccumulationRegisters'
    'РегистрСведений'        = 'InformationRegisters'
    'РегистрРасчета'         = 'CalculationRegisters'
    'ПланВидовРасчета'       = 'ChartsOfCalculationTypes'
    'Перечисление'           = 'Enums'
    'ПланВидовХарактеристик' = 'ChartsOfCharacteristicTypes'
    'ЖурналДокументов'       = 'DocumentJournals'
    'ПланОбмена'             = 'ExchangePlans'
    'Константа'              = 'Constants'
}
$standardFields = @{
    'Справочник'             = 'Ссылка, Код, Наименование, Родитель, Владелец, ЭтоГруппа, ПометкаУдаления, Предопределенный'
    'Документ'               = 'Ссылка, Номер, Дата, Проведен, ПометкаУдаления'
    'РегистрНакопления'      = 'Период, Регистратор, НомерСтроки, Активность, ВидДвижения (для регистров остатков)'
    'РегистрСведений'        = 'Период (если периодический), Регистратор и НомерСтроки (если подчинён регистратору), Активность'
    'РегистрРасчета'         = 'Регистратор, НомерСтроки, Активность, ВидРасчета, ПериодРегистрации, ПериодДействия, ПериодДействияНачало, ПериодДействияКонец, БазовыйПериодНачало, БазовыйПериодКонец, Сторно'
    'ПланВидовРасчета'       = 'Ссылка, Код, Наименование, ПометкаУдаления, Предопределенный'
    'Перечисление'           = 'Ссылка, Порядок'
    'ПланВидовХарактеристик' = 'Ссылка, Код, Наименование, Родитель, ТипЗначения, ПометкаУдаления, Предопределенный'
    'ЖурналДокументов'       = 'Ссылка, Тип, Номер, Дата, Проведен, ПометкаУдаления'
    'ПланОбмена'             = 'Ссылка, Код, Наименование, ЭтотУзел, ПометкаУдаления'
}
$script:index = $null

function Get-Node($node, [string]$path) {
    $xp = ($path -split '/' | ForEach-Object { "*[local-name()='$_']" }) -join '/'
    return $node.SelectSingleNode($xp)
}
function Get-Nodes($node, [string]$path) {
    $xp = ($path -split '/' | ForEach-Object { "*[local-name()='$_']" }) -join '/'
    return $node.SelectNodes($xp)
}
function Get-Synonym($props) {
    $c = Get-Node $props 'Synonym/item/content'
    if ($c) { return $c.InnerText } return ''
}
function Get-Prop($props, [string]$name) {
    $n = Get-Node $props $name
    if ($n) { return $n.InnerText } return ''
}

function Format-Type($props) {
    $t = Get-Node $props 'Type'
    if (-not $t) { return '' }
    $parts = foreach ($n in $t.ChildNodes) {
        if ($n.NodeType -ne 'Element') { continue }
        switch ($n.LocalName) {
            'Type' {
                switch -Regex ($n.InnerText) {
                    'xs:string$' {
                        $len = Get-Prop $t 'StringQualifiers/Length'
                        if ($len -and $len -ne '0') { "Строка($len)" } else { 'Строка' }
                    }
                    'xs:decimal$' { "Число($(Get-Prop $t 'NumberQualifiers/Digits'),$(Get-Prop $t 'NumberQualifiers/FractionDigits'))" }
                    'xs:dateTime$' { 'Дата' }
                    'xs:boolean$' { 'Булево' }
                    default {
                        $n.InnerText -replace '^\w+:', '' -replace '^CatalogRef\.', 'СправочникСсылка.' -replace '^DocumentRef\.', 'ДокументСсылка.' `
                            -replace '^EnumRef\.', 'ПеречислениеСсылка.' -replace '^ChartOfCharacteristicTypesRef\.', 'ПланВидовХарактеристикСсылка.' `
                            -replace '^ExchangePlanRef\.', 'ПланОбменаСсылка.'
                    }
                }
            }
            'TypeSet' { $n.InnerText -replace '^\w+:', '' }
        }
    }
    return (@($parts) -join ' | ')
}

function Format-Fields($parent, [string]$tag, [string]$title, [string]$indent = '') {
    $nodes = Get-Nodes $parent "ChildObjects/$tag"
    if ($nodes.Count -eq 0) { return '' }
    $lines = foreach ($n in $nodes) {
        $p = Get-Node $n 'Properties'
        $syn = Get-Synonym $p
        $line = "$indent  $(Get-Prop $p 'Name'): $(Format-Type $p)"
        if ($syn) { $line += " — $syn" }
        $line
    }
    return "$indent${title}:`n" + (@($lines) -join "`n") + "`n"
}

# --- Метаданные базы через COM (когда нет выгрузки) ---

function Get-ComItems($coll) {
    $n = [int](Invoke-ComMethod $coll 'Count')
    for ($i = 0; $i -lt $n; $i++) { , (Invoke-ComMethod $coll 'Get' $i) }
}
function Get-ComString($v) { [string](Invoke-ComMethod $script:db 'String' $v) }
function Get-ComFullName($m) {
    ([string](Invoke-ComMethod $m 'FullName')) -replace '^(Справочник|Документ|Перечисление|ПланВидовХарактеристик|ПланОбмена|ПланВидовРасчета|ПланСчетов|БизнесПроцесс|Задача)\.', '$1Ссылка.'
}

function Format-ComType($td) {
    $md = Get-ComProp $script:db 'Metadata'
    $types = @(Get-ComItems (Invoke-ComMethod $td 'Types'))
    # Без Select-Object: он оборачивает COM-объект в PSObject, и 1С отвечает «Несоответствие типов»
    $parts = for ($i = 0; $i -lt [Math]::Min($types.Count, 10); $i++) {
        $t = $types[$i]
        $m = Invoke-ComMethod $md 'FindByType' $t
        if ($null -ne $m) { Get-ComFullName $m; continue }
        $s = Get-ComString $t
        switch ($s) {
            'Строка' {
                $len = Get-ComProp (Get-ComProp $td 'StringQualifiers') 'Length'
                if ($len) { "Строка($len)" } else { 'Строка' }
            }
            'Число' {
                $nq = Get-ComProp $td 'NumberQualifiers'
                "Число($(Get-ComProp $nq 'Digits'),$(Get-ComProp $nq 'FractionDigits'))"
            }
            default { $s }
        }
    }
    $res = @($parts) -join ' | '
    if ($types.Count -gt 10) { $res += " | ... (ещё $($types.Count - 10))" }
    return $res
}

function Format-ComFields($obj, [string]$coll, [string]$title, [string]$indent = '') {
    try { $items = @(Get-ComItems (Get-ComProp $obj $coll)) } catch { return '' }
    if (-not $items.Count) { return '' }
    $lines = foreach ($f in $items) {
        $line = "$indent  $(Get-ComProp $f 'Name'): $(Format-ComType (Get-ComProp $f 'Type'))"
        $syn = Get-ComProp $f 'Synonym'
        if ($syn) { $line += " — $syn" }
        $line
    }
    return "$indent${title}:`n" + (@($lines) -join "`n") + "`n"
}

function Get-ComDescribe([string]$kind, [string]$name) {
    Connect-Db
    $m = Invoke-ComMethod (Get-ComProp $script:db 'Metadata') 'FindByFullName' "$kind.$name"
    if ($null -eq $m) { throw "Объект $kind.$name не найден в конфигурации" }
    $sb = New-Object Text.StringBuilder
    [void]$sb.AppendLine("$kind.$name — $(Get-ComProp $m 'Synonym')")
    switch ($kind) {
        'Справочник' {
            [void]$sb.AppendLine("Иерархический: $(Get-ComProp $m 'Hierarchical'); длина кода: $(Get-ComProp $m 'CodeLength'); длина наименования: $(Get-ComProp $m 'DescriptionLength')")
            $owners = @(foreach ($o in Get-ComItems (Get-ComProp $m 'Owners')) { Invoke-ComMethod $o 'FullName' })
            if ($owners.Count) { [void]$sb.AppendLine("Владельцы: $($owners -join ', ')") }
        }
        'РегистрНакопления' {
            $balance = (Get-ComString (Get-ComProp $m 'RegisterType')) -match 'Остат|Balance'
            $vt = if ($balance) { 'Остатки(&Период, Условие), Обороты(&Начало, &Конец, Периодичность, Условие), ОстаткиИОбороты(...)' } else { 'Обороты(&Начало, &Конец, Периодичность, Условие)' }
            [void]$sb.AppendLine("Вид регистра: $(if ($balance) { 'остатки' } else { 'обороты' })")
            [void]$sb.AppendLine("Виртуальные таблицы: $vt")
        }
        'РегистрСведений' {
            $per = Get-ComString (Get-ComProp $m 'InformationRegisterPeriodicity')
            [void]$sb.AppendLine("Периодичность: $per; режим записи: $(Get-ComString (Get-ComProp $m 'WriteMode'))")
            if ($per -notmatch 'Непериод|Nonperiod') { [void]$sb.AppendLine('Виртуальные таблицы: СрезПоследних(&Период, Условие), СрезПервых(&Период, Условие)') }
        }
        'РегистрРасчета' {
            $plan = Get-ComProp $m 'ChartOfCalculationTypes'
            if ($null -ne $plan) { [void]$sb.AppendLine("План видов расчета: $(Invoke-ComMethod $plan 'FullName')") }
            [void]$sb.AppendLine("Периодичность: $(Get-ComString (Get-ComProp $m 'Periodicity')); период действия: $(Get-ComProp $m 'ActionPeriod'); базовый период: $(Get-ComProp $m 'BasePeriod')")
            [void]$sb.AppendLine("Виртуальные таблицы: ФактическийПериодДействия(Условие), ДанныеГрафика(Условие) (если задан график), База<РегистрРасчета>(...)")
        }
        'Константа' { [void]$sb.AppendLine("Тип: $(Format-ComType (Get-ComProp $m 'Type'))") }
    }
    if ($standardFields.Contains($kind)) { [void]$sb.AppendLine("Стандартные поля: $($standardFields[$kind])") }
    [void]$sb.Append((Format-ComFields $m 'Dimensions' 'Измерения'))
    [void]$sb.Append((Format-ComFields $m 'Resources' 'Ресурсы'))
    [void]$sb.Append((Format-ComFields $m 'Attributes' 'Реквизиты'))
    if ($kind -eq 'Перечисление') {
        [void]$sb.AppendLine('Значения (в запросе: ЗНАЧЕНИЕ(Перечисление.' + $name + '.Имя)):')
        foreach ($v in Get-ComItems (Get-ComProp $m 'EnumValues')) { [void]$sb.AppendLine("  $(Get-ComProp $v 'Name') — $(Get-ComProp $v 'Synonym')") }
    }
    try { $tss = @(Get-ComItems (Get-ComProp $m 'TabularSections')) } catch { $tss = @() }
    foreach ($ts in $tss) {
        $tn = Get-ComProp $ts 'Name'
        [void]$sb.AppendLine("Табличная часть $tn — $(Get-ComProp $ts 'Synonym') (таблица запроса: $kind.$name.$tn; поля Ссылка, НомерСтроки)")
        [void]$sb.Append((Format-ComFields $ts 'Attributes' 'Реквизиты' '  '))
    }
    return $sb.ToString().TrimEnd()
}

# --- Метаданные из XML-выгрузки ---

function Get-Index {
    if ($null -ne $script:index) { return $script:index }
    $list = New-Object Collections.ArrayList
    if (-not $cfg.dump) {
        Connect-Db
        $md = Get-ComProp $script:db 'Metadata'
        foreach ($kind in $kinds.Keys) {
            foreach ($o in Get-ComItems (Get-ComProp $md $kinds[$kind])) {
                [void]$list.Add([pscustomobject]@{ Kind = $kind; Name = [string](Get-ComProp $o 'Name'); Synonym = [string](Get-ComProp $o 'Synonym') })
            }
        }
        $script:index = $list
        return $list
    }
    foreach ($kind in $kinds.Keys) {
        $dir = Join-Path $cfg.dump $kinds[$kind]
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        foreach ($f in Get-ChildItem -LiteralPath $dir -Filter *.xml -File) {
            $syn = ''
            # Синоним лежит в начале файла — не грузим весь XML
            $head = [IO.File]::ReadAllText($f.FullName, $utf8)
            $m = [regex]::Match($head, '<Properties>\s*<Name>[^<]*</Name>\s*<Synonym>\s*<v8:item>\s*<v8:lang>[^<]*</v8:lang>\s*<v8:content>([^<]*)</v8:content>')
            if ($m.Success) { $syn = $m.Groups[1].Value }
            [void]$list.Add([pscustomobject]@{ Kind = $kind; Name = $f.BaseName; Synonym = $syn })
        }
    }
    $script:index = $list
    return $list
}

function Resolve-Kind([string]$kind) {
    if ($kinds.Contains($kind)) { return $kind }
    $k = @($kinds.Keys) | Where-Object { $_ -ieq $kind -or $kinds[$_] -ieq $kind -or $kinds[$_] -ieq "${kind}s" }
    if ($k) { return @($k)[0] }
    throw "Неизвестный вид объекта '$kind'. Допустимо: $(@($kinds.Keys) -join ', ')"
}

function Get-Describe($a) {
    if ($a.object) {
        $kind, $name = ([string]$a.object) -split '\.', 2
        $kind = Resolve-Kind $kind
        if (-not $cfg.dump) { return Get-ComDescribe $kind $name }
        $file = Join-Path (Join-Path $cfg.dump $kinds[$kind]) "$name.xml"
        if (-not (Test-Path -LiteralPath $file)) { throw "Объект $kind.$name не найден в выгрузке" }
        $xml = New-Object Xml.XmlDocument
        $xml.Load($file)
        $obj = $xml.DocumentElement.SelectSingleNode('*[1]')
        $props = Get-Node $obj 'Properties'
        $sb = New-Object Text.StringBuilder
        [void]$sb.AppendLine("$kind.$name — $(Get-Synonym $props)")
        switch ($kind) {
            'Справочник' {
                [void]$sb.AppendLine("Иерархический: $(Get-Prop $props 'Hierarchical'); длина кода: $(Get-Prop $props 'CodeLength'); длина наименования: $(Get-Prop $props 'DescriptionLength')")
                $owners = Get-Nodes $props 'Owners/Item'
                if ($owners.Count) { [void]$sb.AppendLine("Владельцы: $(@($owners | ForEach-Object { $_.InnerText -replace '^\w+:', '' }) -join ', ')") }
            }
            'РегистрНакопления' {
                $rt = Get-Prop $props 'RegisterType'
                $vt = if ($rt -eq 'Balance') { 'Остатки(&Период, Условие), Обороты(&Начало, &Конец, Периодичность, Условие), ОстаткиИОбороты(...)' } else { 'Обороты(&Начало, &Конец, Периодичность, Условие)' }
                [void]$sb.AppendLine("Вид регистра: $(if ($rt -eq 'Balance') { 'остатки' } else { 'обороты' })")
                [void]$sb.AppendLine("Виртуальные таблицы: $vt")
            }
            'РегистрСведений' {
                $per = Get-Prop $props 'InformationRegisterPeriodicity'
                [void]$sb.AppendLine("Периодичность: $per; режим записи: $(Get-Prop $props 'WriteMode')")
                if ($per -and $per -ne 'Nonperiodical') { [void]$sb.AppendLine('Виртуальные таблицы: СрезПоследних(&Период, Условие), СрезПервых(&Период, Условие)') }
            }
            'Константа' { [void]$sb.AppendLine("Тип: $(Format-Type $props)") }
        }
        if ($standardFields.Contains($kind)) { [void]$sb.AppendLine("Стандартные поля: $($standardFields[$kind])") }
        [void]$sb.Append((Format-Fields $obj 'Dimension' 'Измерения'))
        [void]$sb.Append((Format-Fields $obj 'Resource' 'Ресурсы'))
        [void]$sb.Append((Format-Fields $obj 'Attribute' 'Реквизиты'))
        $values = Get-Nodes $obj 'ChildObjects/EnumValue'
        if ($values.Count) {
            [void]$sb.AppendLine('Значения (в запросе: ЗНАЧЕНИЕ(Перечисление.' + $name + '.Имя)):')
            foreach ($v in $values) { $p = Get-Node $v 'Properties'; [void]$sb.AppendLine("  $(Get-Prop $p 'Name') — $(Get-Synonym $p)") }
        }
        foreach ($ts in Get-Nodes $obj 'ChildObjects/TabularSection') {
            $p = Get-Node $ts 'Properties'
            [void]$sb.AppendLine("Табличная часть $(Get-Prop $p 'Name') — $(Get-Synonym $p) (таблица запроса: $kind.$name.$(Get-Prop $p 'Name'); поля Ссылка, НомерСтроки)")
            [void]$sb.Append((Format-Fields $ts 'Attribute' 'Реквизиты' '  '))
        }
        return $sb.ToString().TrimEnd()
    }

    $idx = Get-Index
    if ($a.find) {
        $f = [string]$a.find
        $hits = @($idx | Where-Object { $_.Name -like "*$f*" -or $_.Synonym -like "*$f*" })
        if (-not $hits.Count) { return "Ничего не найдено по '$f'" }
        return (@($hits | ForEach-Object { "$($_.Kind).$($_.Name) — $($_.Synonym)" }) -join "`n")
    }
    if ($a.type) {
        $kind = Resolve-Kind ([string]$a.type)
        $hits = @($idx | Where-Object { $_.Kind -eq $kind })
        if ($a.filter) { $f = [string]$a.filter; $hits = @($hits | Where-Object { $_.Name -like "*$f*" -or $_.Synonym -like "*$f*" }) }
        return (@($hits | ForEach-Object { "$($_.Kind).$($_.Name) — $($_.Synonym)" }) -join "`n")
    }
    $lines = foreach ($kind in $kinds.Keys) { "${kind}: $(@($idx | Where-Object { $_.Kind -eq $kind }).Count)" }
    $src = if ($cfg.dump) { "по выгрузке $($cfg.dump)" } else { 'по метаданным базы' }
    return "Конфигурация ($src):`n" + (@($lines) -join "`n")
}

#endregion

#region MCP

$tools = @(
    @{
        name        = 'info'
        description = 'Подключиться к базе 1С и показать конфигурацию, версию и пользователя. Полезно вызвать первым.'
        inputSchema = @{ type = 'object'; properties = @{} }
    },
    @{
        name        = 'describe'
        description = 'Структура метаданных конфигурации. Без параметров — виды объектов с количеством; find — поиск по имени/синониму во всех видах; type (+filter) — список объектов вида; object — поля, табличные части и виртуальные таблицы объекта. Вызывай перед query, чтобы не выдумывать имена полей.'
        inputSchema = @{
            type       = 'object'
            properties = @{
                find   = @{ type = 'string'; description = 'Подстрока имени или синонима, например "контрагент"' }
                type   = @{ type = 'string'; description = 'Вид: Справочник, Документ, РегистрНакопления, РегистрСведений, РегистрРасчета, ПланВидовРасчета, Перечисление, ПланВидовХарактеристик, ЖурналДокументов, ПланОбмена, Константа' }
                filter = @{ type = 'string'; description = 'Подстрока для сужения списка вместе с type' }
                object = @{ type = 'string'; description = 'Полное имя, например "Справочник.Номенклатура" или "РегистрНакопления.ТоварыВРознице"' }
            }
        }
    },
    @{
        name        = 'query'
        description = 'Выполнить запрос на языке запросов 1С (только чтение) и вернуть таблицу: первая строка — колонки, разделитель — табуляция. Ссылки возвращаются представлением. ' + $queryNote + ' Для отбора по ссылкам используй ЗНАЧЕНИЕ(...) или условие по Наименованию/Коду.'
        inputSchema = @{
            type       = 'object'
            properties = @{
                text   = @{ type = 'string'; description = 'Текст запроса, например: ВЫБРАТЬ ПЕРВЫЕ 10 Код, Наименование ИЗ Справочник.Номенклатура' }
                params = @{ type = 'object'; description = 'Параметры запроса {"Имя": значение}. Даты строкой "2026-01-31" или "2026-01-31T23:59:59", массивы — JSON-массивом.' }
                limit  = @{ type = 'integer'; description = 'Сколько строк вернуть (по умолчанию 100, максимум 5000)' }
            }
            required   = @('text')
        }
    }
)

function Send-Message($obj) { $stdout.WriteLine(($obj | ConvertTo-Json -Depth 30 -Compress)) }

function Invoke-Tool([string]$name, $a) {
    if ($null -eq $a) { $a = [pscustomobject]@{} }
    switch ($name) {
        'info' { return Get-DbInfo }
        'describe' { return Get-Describe $a }
        'query' { return Invoke-1CQuery $a }
        default { throw "Неизвестный инструмент: $name" }
    }
}

Write-Log "Запущен, конфиг: $Config"
while ($null -ne ($line = $stdin.ReadLine())) {
    if (-not $line.Trim()) { continue }
    try { $req = $line | ConvertFrom-Json } catch { Write-Log "Некорректный JSON: $line"; continue }
    $hasId = $req.PSObject.Properties.Name -contains 'id'
    if (-not $hasId) { continue }   # уведомления (notifications/*) не требуют ответа
    $resp = @{ jsonrpc = '2.0'; id = $req.id }
    switch ($req.method) {
        'initialize' {
            $pv = if ($req.params.protocolVersion) { $req.params.protocolVersion } else { '2025-06-18' }
            $resp.result = @{
                protocolVersion = $pv
                capabilities    = @{ tools = @{} }
                serverInfo      = @{ name = $serverName; version = '0.2.0' }
                instructions    = $instructions
            }
        }
        'ping' { $resp.result = @{} }
        'tools/list' { $resp.result = @{ tools = $tools } }
        'tools/call' {
            try {
                $text = Invoke-Tool $req.params.name $req.params.arguments
                $resp.result = @{ content = @(@{ type = 'text'; text = [string]$text }); isError = $false }
            } catch {
                $ex = $_.Exception
                while ($ex.InnerException) { $ex = $ex.InnerException }
                $msg = $ex.Message
                Write-Log "Ошибка $($req.params.name): $msg`n$($_.ScriptStackTrace)"
                if ($ex -is [Runtime.InteropServices.COMException] -and $msg -match 'соединени|connection|RPC') { $script:db = $null }
                $resp.result = @{ content = @(@{ type = 'text'; text = "Ошибка: $msg" }); isError = $true }
            }
        }
        default { $resp.error = @{ code = -32601; message = "Method not found: $($req.method)" } }
    }
    Send-Message $resp
}

#endregion
