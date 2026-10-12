# End-to-end tests: builds the connector, starts a fake DSM (mock-dsm.ps1) on
# http://localhost:8765 and evaluates queries with PQTest from the Power Query SDK.
# Exit code 0 when every case matches its expectation.
param([string]$SdkTools = (Join-Path $env:LOCALAPPDATA 'PowerQuerySdkTools\2.155.2'))

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$work = Join-Path $here 'out'
New-Item -ItemType Directory -Force $work | Out-Null
$pqTest = Join-Path $SdkTools 'PQTest.exe'
if (-not (Test-Path $pqTest)) { throw "PQTest.exe not found in $SdkTools (NuGet package Microsoft.PowerQuery.SdkTools)" }

& (Join-Path $here '..\build.ps1') | Out-Null
$mez = Join-Path $here '..\bin\SynologyOffice.mez'
& (Join-Path $here 'make-xlsx.ps1') -Out (Join-Path $work 'report.xlsx')
$log = Join-Path $work 'requests.log'
if (Test-Path $log) { [IO.File]::Delete($log) }

$server = Start-Process powershell -WindowStyle Hidden -PassThru -ArgumentList @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $here 'mock-dsm.ps1'),
    '-Xlsx', (Join-Path $work 'report.xlsx'), '-Log', $log)
Start-Sleep -Seconds 2

$base = 'http://localhost:8765'
$cases = @(
    @{ Name = 'navigator';     Expect = 'Passed'; Contains = '"Name":"Продажи"'; Query = "SynologyOffice.Contents(""$base/oo/r/AbC123#tid=1"")" }
    @{ Name = 'sheet-data';    Expect = 'Passed'; Contains = '"Сумма":250.5';    Query = "SynologyOffice.Contents(""$base/oo/r/AbC123""){[Item=""Продажи"",Kind=""Sheet""]}[Data]" }
    @{ Name = 'hidden-sheet';  Expect = 'Passed'; Contains = '"Value":1';        Query = "Table.RowCount(SynologyOffice.Contents(""$base/oo/r/AbC123""))" }
    @{ Name = 'drive-link';    Expect = 'Passed'; Contains = 'Продажи';          Query = "SynologyOffice.Contents(""$base/d/f/AbC123"")[Name]" }
    @{ Name = 'file-id';       Expect = 'Passed'; Contains = 'Продажи';          Query = "SynologyOffice.Contents(""$base/?file_id=777"")[Name]" }
    @{ Name = 'drive-path';    Expect = 'Passed'; Contains = 'Продажи';          Query = "SynologyOffice.Contents(""$base/?path=/mydrive/Отчёт.osheet"")[Name]" }
    @{ Name = 'unknown-link';  Expect = 'Failed'; Contains = '1002';             Query = "SynologyOffice.Contents(""$base/oo/r/NOPE"")" }
    @{ Name = 'not-osheet';    Expect = 'Failed'; Contains = '.osheet';          Query = "SynologyOffice.Contents(""$base/oo/r/DOC1"")" }
    @{ Name = 'bad-url';       Expect = 'Failed'; Contains = 'oo/r/';             Query = "SynologyOffice.Contents(""$base/something"")" }
    @{ Name = 'https-required'; Expect = 'Failed'; Contains = 'HTTPS';           Query = "SynologyOffice.Contents(""http://nas.example.com/oo/r/AbC123"")" }
    @{ Name = 'export-http-500'; Expect = 'Failed'; Contains = 'HTTP 500';       Query = "SynologyOffice.Contents(""$base/oo/r/BROKEN"")" }
    @{ Name = 'two-loads';     Expect = 'Passed'; Contains = '"Value":4';        Query = "let a = SynologyOffice.Contents(""$base/oo/r/AbC123#tid=1""){0}[Data], b = SynologyOffice.Contents(""$base/oo/r/AbC123#tid=2""){0}[Data] in Table.RowCount(a) + Table.RowCount(b)" }
    @{ Name = 'bad-password';  Expect = 'Failed'; Contains = 'DSM';              Query = "SynologyOffice.Contents(""$base/oo/r/AbC123"")"; User = 'pbi'; Password = 'wrong' }
    @{ Name = '2fa-account';   Expect = 'Failed'; Contains = '2FA';              Query = "SynologyOffice.Contents(""$base/oo/r/AbC123"")"; User = 'otp'; Password = 'x' }
)

# Credentials are keyed by the NAS base URL now, so they are set against the
# host-level data source (SynologyOffice.Database), not against each table query.
# One credential then covers every SynologyOffice.Contents(fullUrl) call on that base.
$credFile = Join-Path $work 'base-cred.query.pq'
[IO.File]::WriteAllText($credFile, "SynologyOffice.Database(""$base"")", (New-Object Text.UTF8Encoding $true))

$failed = 0
try {
    foreach ($c in $cases) {
        $file = Join-Path $work "$($c.Name).query.pq"
        [IO.File]::WriteAllText($file, $c.Query, (New-Object Text.UTF8Encoding $true))
        $user = if ($c.User) { $c.User } else { 'pbi' }
        $password = if ($c.Password) { $c.Password } else { 'secret' }
        @{ AuthenticationKind = 'UsernamePassword'; AuthenticationProperties = @{ Username = $user; Password = $password }
           PrivacySetting = 'None'; Permissions = @() } | ConvertTo-Json -Compress |
            & $pqTest set-credential -e $mez -q $credFile | Out-Null
        $result = ((& $pqTest run-test -e $mez -q $file 2>&1 | Out-String) | ConvertFrom-Json)[0]
        $text = if ($result.Status -eq 'Passed') { $result.Output | ConvertTo-Json -Depth 6 -Compress } else { $result.Error.Message }
        $ok = $result.Status -eq $c.Expect -and $text.Contains($c.Contains)
        if (-not $ok) { $failed++ }
        Write-Host ("{0} {1,-14} {2}: {3}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $c.Name, $result.Status, $text)
    }
    # The auth-error cases above left bad credentials stored for this data source path.
    '{"AuthenticationKind":"UsernamePassword","AuthenticationProperties":{"Username":"pbi","Password":"secret"},"PrivacySetting":"None","Permissions":[]}' |
        & $pqTest set-credential -e $mez -q $credFile | Out-Null
    $tc = (& $pqTest test-connection -e $mez -q $credFile 2>&1 | Out-String)
    $ok = $tc -match 'TestConnection succeeded'
    if (-not $ok) { $failed++ }
    Write-Host ("{0} test-connection" -f $(if ($ok) { 'PASS' } else { 'FAIL' }))

    $lines = Get-Content $log -Encoding UTF8
    $check = {
        param([bool]$ok, [string]$name)
        if (-not $ok) { $script:failed++ }
        Write-Host ("{0} {1}" -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $name)
    }
    & $check (-not ($lines -match 'auth=\[.+\]')) 'no Authorization header sent'
    & $check (-not ($lines -match 'query=\[[^\]]*_sid')) 'sid never in URL'
    # Every session that downloaded an export must be logged out afterwards.
    $sidOf = { param($line) if ($line -match '(?:^|&| )_sid=([^& ]+)') { $Matches[1] } }
    $downloads = @(); $logouts = @{}
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match 'api=SYNO\.Office\.Export') { $downloads += , @($i, (& $sidOf $lines[$i])) }
        if ($lines[$i] -match 'method=logout') { $logouts[(& $sidOf $lines[$i])] = $i }
    }
    $orphans = @($downloads | Where-Object { -not $logouts.ContainsKey($_[1]) -or $logouts[$_[1]] -lt $_[0] })
    & $check ($downloads.Count -gt 0 -and $orphans.Count -eq 0) "logout after each export ($($downloads.Count) exports)"
} finally {
    try { Invoke-WebRequest "$base/quit" -UseBasicParsing -TimeoutSec 5 | Out-Null } catch { }
    if (-not $server.HasExited) { Stop-Process -Id $server.Id -Force }
}

Write-Host "$failed failed"
exit $failed
