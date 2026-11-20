param([string]$Config, [string]$QueryFile, [string]$OutFile, [string]$ParamsJson = '{}')
$ErrorActionPreference = 'Stop'
$inv = [Globalization.CultureInfo]::InvariantCulture
$cfg = Get-Content $Config -Raw -Encoding UTF8 | ConvertFrom-Json
$comType = [System.__ComObject]
function Get-ComProp($o, [string]$name) { , $comType.InvokeMember($name, 'GetProperty', $null, $o, $null) }
function Set-ComProp($o, [string]$name, $value) { [void]$comType.InvokeMember($name, 'SetProperty', $null, $o, @(, $value)) }
function Invoke-ComMethod($o, [string]$name) { , $comType.InvokeMember($name, 'InvokeMethod', $null, $o, [object[]]$args) }
$ss = ConvertTo-SecureString ([string]$cfg.passwordEnc)
$b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
$pwd = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($b)
$conn = 'Srvr="{0}";Ref="{1}";Usr="{2}";Pwd="{3}";' -f $cfg.server, $cfg.ref, $cfg.user, $pwd.Replace('"','""')
$connector = New-Object -ComObject V83.COMConnector
$db = $connector.Connect($conn)
function Fmt($v) {
    if ($null -eq $v -or $v -is [DBNull]) { return '' }
    if ($v -is [datetime]) { if ($v.Year -le 100) { return '' } return $v.ToString('yyyy-MM-dd') }
    if ($v -is [bool]) { if ($v) { return '1' } return '0' }
    if ($v -is [string]) { return ($v -replace '[\t\r\n]+', ' ') }
    if ($v -is [ValueType]) { return [Convert]::ToString($v, $inv) }
    return ([string](Invoke-ComMethod $db 'String' $v)) -replace '[\t\r\n]+', ' '
}
$q = Invoke-ComMethod $db 'NewObject' 'Query'
Set-ComProp $q 'Text' ([IO.File]::ReadAllText($QueryFile, [Text.Encoding]::UTF8))
if (Test-Path $ParamsJson) { $ParamsJson = [IO.File]::ReadAllText($ParamsJson,[Text.Encoding]::UTF8) }; $pr = $ParamsJson | ConvertFrom-Json
foreach ($p in $pr.PSObject.Properties) { [void](Invoke-ComMethod $q 'SetParameter' $p.Name ([datetime]::Parse($p.Value, $inv))) }
$res = Invoke-ComMethod $q 'Execute'
$columns = Get-ComProp $res 'Columns'
$n = Invoke-ComMethod $columns 'Count'
$cols = for ($i = 0; $i -lt $n; $i++) { Get-ComProp (Invoke-ComMethod $columns 'Get' $i) 'Name' }
$sw = New-Object IO.StreamWriter($OutFile, $false, (New-Object Text.UTF8Encoding($false)))
$sw.WriteLine((@($cols) -join "`t"))
$sel = Invoke-ComMethod $res 'Select'
$cnt = 0
while (Invoke-ComMethod $sel 'Next') {
    $vals = for ($i = 0; $i -lt $n; $i++) { Fmt (Invoke-ComMethod $sel 'Get' $i) }
    $sw.WriteLine((@($vals) -join "`t")); $cnt++
}
$sw.Close()
Write-Output "$OutFile : $cnt rows"

