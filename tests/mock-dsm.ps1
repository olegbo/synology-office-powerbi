# Minimal fake DSM Web API on http://localhost:8765/ for connector tests.
# Every login gets a new sid that stops working after logout, like real DSM.
# Logs one line per request: method, path, raw URL query and all parameters (except passwd).
param([string]$Xlsx, [string]$Log)
Add-Type -AssemblyName System.Web
$l = New-Object Net.HttpListener
$l.Prefixes.Add('http://localhost:8765/')
$l.Start()
$sessions = @{}
$nextSid = 1
function Send($ctx, [byte[]]$bytes, $type) {
    $ctx.Response.ContentType = $type
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $ctx.Response.Close()
}
function Json($ctx, $s) { Send $ctx ([Text.Encoding]::UTF8.GetBytes($s)) 'application/json' }
while ($l.IsListening) {
    $ctx = $l.GetContext()
    $req = $ctx.Request
    $body = (New-Object IO.StreamReader $req.InputStream).ReadToEnd()
    $q = [Web.HttpUtility]::ParseQueryString($req.Url.Query)
    $b = [Web.HttpUtility]::ParseQueryString($body)
    $p = @{}; foreach ($c in $q, $b) { foreach ($k in $c.Keys) { if ($k) { $p[$k] = $c[$k] } } }
    $logged = ($p.GetEnumerator() | Sort-Object Key | ? Key -ne 'passwd' | % { "$($_.Key)=$($_.Value)" }) -join '&'
    Add-Content $Log "$($req.HttpMethod) $($req.Url.AbsolutePath) query=[$($req.Url.Query)] auth=[$($req.Headers['Authorization'])] $logged" -Encoding UTF8
    if ($req.Url.AbsolutePath -eq '/quit') { Json $ctx '{}'; break }
    $valid = $p['_sid'] -and $sessions.ContainsKey($p['_sid'])
    switch ($p['api']) {
        'SYNO.API.Auth' {
            if ($p['method'] -eq 'logout') {
                if ($valid) { $sessions.Remove($p['_sid']); Json $ctx '{"success":true}' } else { Json $ctx '{"success":false,"error":{"code":119}}' }
            }
            elseif ($p['account'] -eq 'pbi' -and $p['passwd'] -eq 'secret') {
                $sid = "SID$nextSid"; $nextSid++; $sessions[$sid] = $true
                Json $ctx "{""success"":true,""data"":{""sid"":""$sid""}}"
            }
            elseif ($p['account'] -eq 'otp') { Json $ctx '{"success":false,"error":{"code":403}}' }
            else { Json $ctx '{"success":false,"error":{"code":400}}' }
        }
        'SYNO.SynologyDrive.Files' {
            if (-not $valid) { Json $ctx '{"success":false,"error":{"code":119}}' }
            elseif ($p['path'] -in 'link:AbC123', 'id:777', '/mydrive/Отчёт.osheet') { Json $ctx '{"success":true,"data":{"name":"Отчёт.osheet","file_id":"777"}}' }
            elseif ($p['path'] -eq 'link:BROKEN') { Json $ctx '{"success":true,"data":{"name":"Broken.osheet","file_id":"888"}}' }
            elseif ($p['path'] -eq 'link:DOC1') { Json $ctx '{"success":true,"data":{"name":"Notes.odoc","file_id":"5"}}' }
            else { Json $ctx '{"success":false,"error":{"code":1002}}' }
        }
        'SYNO.Office.Export' {
            if (-not $valid) { Json $ctx '{"success":false,"error":{"code":119}}' }
            elseif ($p['path'] -eq 'id:777') { Send $ctx ([IO.File]::ReadAllBytes($Xlsx)) 'application/octet-stream' }
            elseif ($p['path'] -eq 'id:888') { $ctx.Response.StatusCode = 500; Send $ctx ([Text.Encoding]::UTF8.GetBytes('<html>Internal Server Error</html>')) 'text/html' }
            else { Json $ctx '{"success":false,"error":{"code":1001}}' }
        }
        default { $ctx.Response.StatusCode = 404; Json $ctx '<html>nope</html>' }
    }
}
$l.Stop()
