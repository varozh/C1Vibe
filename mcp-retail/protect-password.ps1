# Шифрует пароль в config.json MCP-сервера 1С через Windows DPAPI (область текущего пользователя).
# Открытое поле "password" заменяется на "passwordEnc"; расшифровать его может только этот же пользователь
# Windows на этом же компьютере, поэтому файл можно класть в git.
# Если в конфиге нет открытого пароля (или указан -Prompt), пароль спрашивается с клавиатуры.
# На другом компьютере / под другим пользователем запустите скрипт заново с -Prompt.
#   powershell -ExecutionPolicy Bypass -File protect-password.ps1 -Config ..\mcp-trade\config.json
param(
    [Parameter(Mandatory = $true)][string]$Config,
    [switch]$Prompt
)
$ErrorActionPreference = 'Stop'
$utf8 = New-Object Text.UTF8Encoding($false)
$path = (Resolve-Path -LiteralPath $Config).Path
$text = [IO.File]::ReadAllText($path, $utf8)
$cfg = $text | ConvertFrom-Json

if (-not $Prompt -and $null -ne $cfg.password) {
    $ss = ConvertTo-SecureString ([string]$cfg.password) -AsPlainText -Force
} else {
    $ss = Read-Host "Пароль пользователя 1С '$($cfg.user)'" -AsSecureString
}
$enc = ConvertFrom-SecureString $ss

$pwdLine = '"password"\s*:\s*"(?:[^"\\]|\\.)*"'
$encLine = '"passwordEnc"\s*:\s*"[^"]*"'
if ($text -match $encLine) {
    $text = [regex]::Replace($text, $encLine, "`"passwordEnc`": `"$enc`"")
    $text = [regex]::Replace($text, '\s*' + $pwdLine + '\s*,?', '')
} elseif ($text -match $pwdLine) {
    $text = [regex]::Replace($text, $pwdLine, "`"passwordEnc`": `"$enc`"")
} else {
    $text = [regex]::Replace($text, '"user"\s*:\s*"(?:[^"\\]|\\.)*"', "`$0,`n  `"passwordEnc`": `"$enc`"")
}
[void]($text | ConvertFrom-Json)   # проверка, что JSON остался корректным
[IO.File]::WriteAllText($path, $text, $utf8)
Write-Host "Пароль зашифрован: $path"
