# Бот MAX: напоминание магазинам об обмене с центральной базой Розницы и утренний список тех, кто обмен не выполнил.
# Разрядность PowerShell должна совпадать с разрядностью зарегистрированного COM-коннектора 1С (comcntr.dll):
#   32 бита — C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe
#   64 бита — C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
# Режимы:
#   Reminder     — вечернее напоминание всем выполнить обмен (база не нужна);
#   Report       — список магазинов без успешной загрузки в центральной базе после eveningFrom прошлого дня;
#   ListChats    — ждать события 3 минуты и показать chat_id групп, куда добавили бота;
#   SetToken     — сохранить токен бота в config.json в зашифрованном виде (DPAPI);
#   SetPassword  — сохранить пароль пользователя 1С в config.json в зашифрованном виде (DPAPI);
#   Check        — проверить окружение: COM-коннектор, подключение к 1С, доступ к MAX, chatId;
#   InstallTasks — создать задачи в Планировщике (Reminder и Report) от текущего пользователя
#                  тем же powershell.exe, которым запущен скрипт.
# -DryRun: ничего не отправлять, только вывести текст сообщения.
param(
    [Parameter(Mandatory)][ValidateSet('Reminder', 'Report', 'ListChats', 'SetToken', 'SetPassword', 'Check', 'InstallTasks')][string]$Mode,
    [string]$Config,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Config) { $Config = Join-Path $scriptDir 'config.json' }
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$cfg = Get-Content $Config -Raw -Encoding UTF8 | ConvertFrom-Json
$logFile = Join-Path $scriptDir 'exchange-bot.log'

function Write-Log([string]$msg) {
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Mode, $msg
    Write-Host $line
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Unprotect-Text([string]$enc) {
    $ss = ConvertTo-SecureString $enc
    $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
}

#region MAX Bot API

function Get-Token {
    if (-not $cfg.tokenEnc) { throw "Токен бота не задан. Выполните: exchange-bot.ps1 -Mode SetToken" }
    try { return Unprotect-Text ([string]$cfg.tokenEnc) }
    catch { throw "Не удалось расшифровать tokenEnc: он зашифрован другим пользователем Windows или на другом компьютере. Выполните -Mode SetToken заново." }
}

function Invoke-Max([string]$method, [string]$path, $body) {
    $p = @{ Method = $method; Uri = "$($cfg.apiUrl)$path"; Headers = @{ Authorization = (Get-Token) }; UseBasicParsing = $true }
    if ($null -ne $body) {
        $p.ContentType = 'application/json; charset=utf-8'
        $p.Body = [Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Depth 10 -Compress))
    }
    # Windows Server 2012 R2 изредка обрывает TLS-рукопожатие (Schannel 36887, bad_record_mac). Повторяем только
    # ошибки установки соединения — запрос тогда до MAX не дошёл, и сообщение не задвоится. Ответы 4xx/5xx не повторяем.
    $retryable = 'SecureChannelFailure', 'ConnectFailure', 'NameResolutionFailure', 'TrustFailure'
    for ($attempt = 1; ; $attempt++) {
        try { $r = Invoke-WebRequest @p; break }
        catch {
            $we = $_.Exception -as [Net.WebException]
            if ($attempt -ge 5 -or -not $we -or $retryable -notcontains [string]$we.Status) { throw }
            Write-Log "Нет соединения с MAX ($($we.Status)), попытка $attempt из 5, повтор через 5 с"
            Start-Sleep -Seconds 5
        }
    }
    # Invoke-WebRequest в PowerShell 5.1 может неверно угадать кодировку ответа — читаем байты как UTF-8
    $ms = New-Object IO.MemoryStream
    $r.RawContentStream.Position = 0
    $r.RawContentStream.CopyTo($ms)
    return [Text.Encoding]::UTF8.GetString($ms.ToArray()) | ConvertFrom-Json
}

function Send-MaxText([string]$text) {
    if ($DryRun) { Write-Host "---- DryRun, сообщение не отправлено ----`n$text`n-----------------------------------------"; return }
    if (-not $cfg.chatId) { throw "Не задан chatId в config.json. Узнайте его: exchange-bot.ps1 -Mode ListChats" }
    # Лимит MAX — 4000 символов на сообщение: длинный текст режем по строкам
    $parts = New-Object Collections.Generic.List[string]
    $cur = ''
    foreach ($line in $text -split "`n") {
        if (($cur.Length + $line.Length + 1) -gt 3900 -and $cur) { $parts.Add($cur); $cur = '' }
        $cur = if ($cur) { "$cur`n$line" } else { $line }
    }
    if ($cur) { $parts.Add($cur) }
    foreach ($part in $parts) {
        [void](Invoke-Max 'POST' "/messages?chat_id=$($cfg.chatId)" @{ text = $part })
        Start-Sleep -Milliseconds 600
    }
    Write-Log "Отправлено в чат $($cfg.chatId): $($parts.Count) сообщ."
}

#endregion

#region Данные Розницы

$comType = [System.__ComObject]
function Get-ComProp($o, [string]$name) { , $comType.InvokeMember($name, 'GetProperty', $null, $o, $null) }
function Set-ComProp($o, [string]$name, $value) { [void]$comType.InvokeMember($name, 'SetProperty', $null, $o, @(, $value)) }
function Invoke-ComMethod($o, [string]$name) { , $comType.InvokeMember($name, 'InvokeMethod', $null, $o, [object[]]$args) }

# Параметры подключения: раздел "retail" в config.json; если его нет — файл retailConfig (конфиг MCP-сервера).
function Get-RetailSettings {
    if ($cfg.retail -and $cfg.retail.server) { return $cfg.retail }
    if (-not $cfg.retailConfig) { throw 'В config.json не заполнен раздел "retail" (server, ref, user).' }
    $path = [string]$cfg.retailConfig
    if (-not [IO.Path]::IsPathRooted($path)) { $path = Join-Path $scriptDir $path }
    return Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json
}

function New-ComConnector {
    try { return New-Object -ComObject V83.COMConnector }
    catch {
        $other = if ([IntPtr]::Size -eq 8) { '32-битном (C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe)' } else { '64-битном (C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe)' }
        throw "COM-коннектор 1С не зарегистрирован для $([IntPtr]::Size * 8)-битного процесса. Зарегистрируйте comcntr.dll (regsvr32) нужной разрядности или запустите скрипт в $other PowerShell."
    }
}

function Connect-Retail {
    $rc = Get-RetailSettings
    if (-not $rc.passwordEnc -and -not $rc.password) { throw 'Не задан пароль 1С. Выполните: exchange-bot.ps1 -Mode SetPassword' }
    try { $pwd1c = if ($rc.passwordEnc) { Unprotect-Text ([string]$rc.passwordEnc) } else { [string]$rc.password } }
    catch { throw 'Не удалось расшифровать пароль 1С: он зашифрован другим пользователем Windows или на другом компьютере. Выполните -Mode SetPassword заново.' }
    $q = { param($s) ([string]$s).Replace('"', '""') }
    $conn = 'Srvr="{0}";Ref="{1}";Usr="{2}";Pwd="{3}";' -f (& $q $rc.server), (& $q $rc.ref), (& $q $rc.user), (& $q $pwd1c)
    return (New-ComConnector).Connect($conn)
}

function Save-Config($raw) {
    [IO.File]::WriteAllText($Config, ($raw | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
}

# Магазины (узлы плана обмена «ПоМагазину») и дата последней успешной загрузки их данных в центральную базу.
# Это те же данные, что в форме «Выполнение обмена данными» → «Последняя удачная загрузка данных».
$queryText = @'
ВЫБРАТЬ
    Узлы.Код КАК Код,
    Узлы.Наименование КАК Наименование,
    Ист.Дата КАК Дата
ИЗ
    ПланОбмена.ПоМагазину КАК Узлы
    // Внутреннее соединение: магазины без единой успешной загрузки (ещё не открыты) в список не попадают
    ВНУТРЕННЕЕ СОЕДИНЕНИЕ (ВЫБРАТЬ
            Х.УзелИнформационнойБазы КАК Узел,
            МАКСИМУМ(Х.ДатаПоследнейЗагрузки) КАК Дата
        ИЗ
            РегистрСведений.ИсторияОбменаДанными КАК Х
        ГДЕ
            Х.РезультатПоследнейЗагрузки = ИСТИНА
        СГРУППИРОВАТЬ ПО
            Х.УзелИнформационнойБазы) КАК Ист
    ПО Ист.Узел = Узлы.Ссылка
ГДЕ
    НЕ Узлы.ПометкаУдаления
    И (Узлы.Наименование ПОДОБНО "%Business Line%" ИЛИ Узлы.Наименование ПОДОБНО "%BL%" ИЛИ Узлы.Наименование ПОДОБНО "%БЛ%")
    И НЕ Узлы.Наименование ПОДОБНО "%SoLi%"
    И НЕ Узлы.Наименование ПОДОБНО "%Gesolin%"
    И НЕ Узлы.Наименование ПОДОБНО "%ОФИС%"
    И Ист.Дата < &Граница
УПОРЯДОЧИТЬ ПО
    Наименование
'@

function Get-MissedStores([datetime]$border) {
    $db = Connect-Retail
    $q = Invoke-ComMethod $db 'NewObject' 'Query'
    Set-ComProp $q 'Text' $queryText
    [void](Invoke-ComMethod $q 'SetParameter' 'Граница' $border)
    $sel = Invoke-ComMethod (Invoke-ComMethod $q 'Execute') 'Select'
    $exclude = @($cfg.exclude | ForEach-Object { ([string]$_).Trim() })
    $list = while (Invoke-ComMethod $sel 'Next') {
        $code = ([string](Get-ComProp $sel 'Код')).Trim()
        $name = ([string](Get-ComProp $sel 'Наименование')).Trim()
        if ($exclude -contains $code -or $exclude -contains $name) { continue }
        $d = Get-ComProp $sel 'Дата'
        [pscustomobject]@{ Code = $code; Name = $name; Last = [datetime]$d }
    }
    return @($list)
}

#endregion

try {
    switch ($Mode) {
        'SetToken' {
            $ss = Read-Host 'Токен бота MAX' -AsSecureString
            $raw = Get-Content $Config -Raw -Encoding UTF8 | ConvertFrom-Json
            $raw.tokenEnc = ConvertFrom-SecureString $ss
            Save-Config $raw
            Write-Host 'Токен сохранён в config.json (tokenEnc).'
        }
        'SetPassword' {
            $raw = Get-Content $Config -Raw -Encoding UTF8 | ConvertFrom-Json
            if (-not $raw.retail) { throw 'В config.json нет раздела "retail" (server, ref, user).' }
            $ss = Read-Host "Пароль пользователя 1С «$($raw.retail.user)»" -AsSecureString
            $raw.retail.passwordEnc = ConvertFrom-SecureString $ss
            Save-Config $raw
            Write-Host 'Пароль сохранён в config.json (retail.passwordEnc).'
        }
        'Check' {
            $ok = $true
            $step = { param($name, [scriptblock]$body)
                try { $r = & $body; Write-Host "[OK]     $name$(if ($r) { ": $r" })" -ForegroundColor Green }
                catch { Write-Host "[ОШИБКА] ${name}: $($_.Exception.Message)" -ForegroundColor Red; Set-Variable ok $false -Scope 1 }
            }
            & $step 'PowerShell' { "$($PSVersionTable.PSVersion), $([IntPtr]::Size * 8) бит, $([Diagnostics.Process]::GetCurrentProcess().Path)" }
            & $step 'Пользователь Windows' { "$env:USERDOMAIN\$env:USERNAME" }
            & $step 'COM-коннектор 1С' { [void](New-ComConnector); 'V83.COMConnector создан' }
            & $step 'Подключение к Рознице' {
                $rc = Get-RetailSettings
                $db = Connect-Retail
                $q = Invoke-ComMethod $db 'NewObject' 'Query'
                Set-ComProp $q 'Text' 'ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Н ИЗ ПланОбмена.ПоМагазину'
                $sel = Invoke-ComMethod (Invoke-ComMethod $q 'Execute') 'Select'
                [void](Invoke-ComMethod $sel 'Next')
                "$($rc.server)/$($rc.ref), узлов обмена: $(Get-ComProp $sel 'Н')"
            }
            & $step 'Доступ к MAX и токен' { $me = Invoke-Max 'GET' '/me' $null; "бот «$($me.name)»" }
            & $step 'Группа для рассылки' {
                if (-not $cfg.chatId) { throw 'chatId не задан — выполните -Mode ListChats' }
                $c = Invoke-Max 'GET' "/chats/$($cfg.chatId)" $null
                "$($cfg.chatId) «$($c.title)»"
            }
            if ($ok) { Write-Host 'Всё готово к работе.' -ForegroundColor Green } else { exit 1 }
        }
        'InstallTasks' {
            $exe = [Diagnostics.Process]::GetCurrentProcess().Path
            $script = $MyInvocation.MyCommand.Path
            $user = "$env:USERDOMAIN\$env:USERNAME"
            Write-Host "Задачи будут запускаться от $user через $exe"
            Write-Host 'Пароль Windows нужен, чтобы задачи выполнялись без входа пользователя в систему.'
            $ss = Read-Host "Пароль Windows пользователя $user" -AsSecureString
            $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
            try { $winPwd = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) }
            $tasks = @(
                @{ Name = 'MAX обмен магазинов - напоминание'; Mode = 'Reminder'; Time = [string]$cfg.reminderTime },
                @{ Name = 'MAX обмен магазинов - отчет'; Mode = 'Report'; Time = [string]$cfg.reportTime }
            )
            foreach ($t in $tasks) {
                $action = New-ScheduledTaskAction -Execute $exe -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -Mode $($t.Mode)" -WorkingDirectory $scriptDir
                $trigger = New-ScheduledTaskTrigger -Daily -At $t.Time
                $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew
                [void](Register-ScheduledTask -TaskName $t.Name -Action $action -Trigger $trigger -Settings $settings -User $user -Password $winPwd -RunLevel Limited -Force)
                Write-Host "Создана задача «$($t.Name)»: ежедневно в $($t.Time)"
            }
        }
        'ListChats' {
            $me = Invoke-Max 'GET' '/me' $null
            Write-Host "Бот: $($me.name) (@$($me.username))"
            # В MAX нет метода «список чатов бота»: chat_id приходит в событиях (бота добавили в группу, сообщение в группе).
            Write-Host 'Жду события 3 минуты: добавьте бота в группу или напишите в группе сообщение...'
            $seen = @{}
            $marker = $null
            $until = (Get-Date).AddMinutes(3)
            while ((Get-Date) -lt $until) {
                $path = '/updates?timeout=30&limit=100'
                if ($marker) { $path += "&marker=$marker" }
                $r = Invoke-Max 'GET' $path $null
                if ($r.marker) { $marker = $r.marker }
                foreach ($u in @($r.updates)) {
                    $id = if ($u.chat_id) { $u.chat_id } elseif ($u.message.recipient.chat_id) { $u.message.recipient.chat_id } else { $null }
                    if (-not $id -or $seen.ContainsKey([string]$id)) { continue }
                    $seen[[string]$id] = $true
                    $title = ''
                    try { $title = (Invoke-Max 'GET' "/chats/$id" $null).title } catch { }
                    Write-Host ('chat_id = {0}   {1}   (событие {2})' -f $id, $title, $u.update_type)
                }
            }
            if ($seen.Count -eq 0) { Write-Host 'Событий не было. Запустите ListChats снова и, пока идёт ожидание, удалите бота из группы и добавьте обратно (можно обычным участником). Или возьмите id группы из адресной строки web.max.ru.' }
        }
        'Reminder' {
            Send-MaxText ([string]$cfg.reminderText)
        }
        'Report' {
            $from = [timespan]::Parse([string]$cfg.eveningFrom)
            $border = (Get-Date).Date.AddDays(-1).Add($from)
            $missed = @(Get-MissedStores $border)
            Write-Log "Граница $($border.ToString('dd.MM.yyyy HH:mm')), не выполнили: $($missed.Count)"
            if ($missed.Count -eq 0) {
                Send-MaxText ([string]$cfg.reportAllOk)
            } else {
                $lines = @([string]$cfg.reportHeader)
                $lines += $missed | ForEach-Object {
                    "• $($_.Name) (последний обмен $($_.Last.ToString('dd.MM HH:mm')))"
                }
                if ($cfg.reportFooter) { $lines += ''; $lines += [string]$cfg.reportFooter }
                Send-MaxText ($lines -join "`n")
            }
        }
    }
}
catch {
    Write-Log "ОШИБКА: $($_.Exception.Message)"
    exit 1
}
