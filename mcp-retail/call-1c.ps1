# Обёртка для одноразового вызова инструмента MCP-сервера 1С.
# Использование: .\call-1c.ps1 -Tool query -ArgsJson '{"text":"..."}'
#          или:  .\call-1c.ps1 -Tool describe -ArgsJson '{"find":"магазин"}'
param(
    [Parameter(Mandatory)][string]$Tool,
    [string]$ArgsJson = '',
    [string]$ArgsFile = '',
    [string]$Config = 'D:\002_BASE\Vibe\mcp-retail\config.json'
)
$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = New-Object Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
if ($ArgsFile) { $ArgsJson = [IO.File]::ReadAllText($ArgsFile, [Text.Encoding]::UTF8) }
if (-not $ArgsJson) { $ArgsJson = '{}' }
$server = 'D:\002_BASE\Vibe\mcp-retail\server.ps1'
$psi = New-Object Diagnostics.ProcessStartInfo
if ($Config -match 'mcp-hrm') {
    $psi.FileName = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
} else {
    $psi.FileName = 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
}
$psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$server`" -Config `"$Config`""
$psi.RedirectStandardInput = $true
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$psi.StandardOutputEncoding = New-Object Text.UTF8Encoding($false)
$psi.StandardErrorEncoding = New-Object Text.UTF8Encoding($false)

$p = [Diagnostics.Process]::Start($psi)
$stderr = ''
$null = [void]([System.Threading.Tasks.Task]::Run([Action]{
    $line = ''
    while (($line = $p.StandardError.ReadLine()) -ne $null) {
        [Console]::Error.WriteLine("[srv] $line")
    }
}))

$reqs = @(
    @{ jsonrpc='2.0'; id=1; method='initialize'; params=@{ protocolVersion='2025-06-18'; capabilities=@{}; clientInfo=@{name='caller';version='1.0'} } },
    @{ jsonrpc='2.0'; method='notifications/initialized' },
    @{ jsonrpc='2.0'; id=2; method='tools/call'; params=@{ name=$Tool; arguments=($ArgsJson | ConvertFrom-Json) } }
)
foreach ($r in $reqs) { $p.StandardInput.WriteLine(($r | ConvertTo-Json -Depth 30 -Compress)) }

# читаем ответы, пока не получим id=2
$sw = [Diagnostics.Stopwatch]::StartNew()
while ($sw.Elapsed.TotalSeconds -lt 300) {
    if ($p.StandardOutput.EndOfStream) { break }
    $line = $p.StandardOutput.ReadLine()
    if (-not $line) { break }
    $resp = $line | ConvertFrom-Json
    if ($resp.id -eq 2) {
        if ($resp.result.content) { $resp.result.content | ForEach-Object { $_.text } }
        elseif ($resp.error) { "ОШИБКА: $($resp.error.message)" }
        break
    }
}
try { $p.Kill() } catch {}
