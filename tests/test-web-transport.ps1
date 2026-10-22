[CmdletBinding()]
param([switch]$NoNetwork)
# 2026-10-22 김좌현 수정 : 실제 사이트 저장 없이 Excel VBA와 로컬 HTTP 서버로 웹 전송 회귀 검증
# 2026-10-22 김좌현 수정 : 요청번호 수정 폼·웹 전용 내용 보존·충돌·부분 성공 fixture 확장
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$probe = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, 0)
$probe.Start()
$fixturePort = $probe.LocalEndpoint.Port
$probe.Stop()
$serverTrace = Join-Path ([IO.Path]::GetTempPath()) ('web-fixture-' + [guid]::NewGuid().ToString('N') + '.trace')
$fixtureJob = Start-Job -ArgumentList $fixturePort,$serverTrace -ScriptBlock {
    param($port, $serverTrace)
    $ErrorActionPreference = 'Stop'
    $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $port)
    $listener.Start()
    $sequence = 0
    $savePosts = 0
    $updatePosts = 0
    $saved = @{}
    $existingViews = @{}
    function Field([string]$name, [string]$value) {
        return '<input name="' + $name + '" id="' + $name + '" value="' + [Net.WebUtility]::HtmlEncode($value) + '">'
    }
    function Registration([string]$number, [string]$state) {
        return '<form method="post" action="./c0210_in.aspx"><input type="hidden" name="__VIEWSTATE" value="' + [Net.WebUtility]::HtmlEncode($state) + '">' +
            '<span id="REQ_NO">' + $number + '</span>' + (Field 'REQ_TIT' '') + '<textarea name="REQ_RMK" id="REQ_RMK"></textarea>' +
            '<select name="GUBUN" id="GUBUN"><option value="PJT">Project</option></select>' +
            '<select name="REQ_UID" id="REQ_UID"><option value="u1">Same name</option><option value="u2">Same name</option></select>' +
            '<select name="STS" id="STS"><option value="1">In progress</option></select>' +
            (Field 'PGM_NM' '') + (Field 'PGM_ID' '') + (Field 'ACT_TIME' '30') + '<input type="image" name="btnSave"></form>'
    }
    function Choice([string]$name, [string]$selected, [string[]]$options) {
        $html = '<select name="' + $name + '" id="' + $name + '">'
        foreach ($option in $options) {
            $html += '<option value="' + [Net.WebUtility]::HtmlEncode($option) + '"' + $(if ($option -ceq $selected) { ' selected' } else { '' }) + '>' + [Net.WebUtility]::HtmlEncode($option) + '</option>'
        }
        return $html + '</select>'
    }
    function Existing([string]$number, [hashtable]$record, [string]$state) {
        $html = '<form method="post" action="./c0200_02.aspx?REQ_NO=' + $number + '">' + (Field '__VIEWSTATE' $state) +
            (Field 'seqNum' $number) +
            '<span id="REQ_NO">' + $number + '</span><span id="REQ_UID">u1</span>'
        # Real ASP.NET WinHTTP GET omits these hidden fields despite rendering the anchors.
        if ($number -ne 'RQ9101') { $html += (Field '__EVENTTARGET' '') + (Field '__EVENTARGUMENT' '') }
        foreach ($name in @('REQ_TIT','PGM_NM','PGM_ID','ACT_DT_FR','ACT_DT','ACT_TIME','ADM_REMARK','ACCPT_RMK','unique_key')) { $html += Field $name $record[$name] }
        foreach ($name in @('REQ_RMK','ACT_RMK')) { $html += '<textarea name="' + $name + '" id="' + $name + '">' + "`r`n" + [Net.WebUtility]::HtmlEncode($record[$name]) + '</textarea>' }
        $html += Choice 'GUBUN' $record['GUBUN'] @('PJT','MA')
        $html += Choice 'STS' $record['STS'] @('1','5','9')
        $html += Choice 'PGM_MD' $record['PGM_MD'] @('','M')
        $html += Choice 'ACT_GUBUN' $record['ACT_GUBUN'] @('','300')
        $html += Choice 'ACT_UID' $record['ACT_UID'] @('worker','other-worker')
        return $html + '<a id="btnSaveReq" href="javascript:__doPostBack(&#39;btnSaveReq&#39;,&#39;&#39;)"></a>' +
            '<a id="btnSave" href="javascript:__doPostBack(&#39;btnSave&#39;,&#39;&#39;)"></a></form>'
    }
    foreach ($number in @('RQ9101','RQ9102','RQ9103','RQ9104','RQ9199')) {
        $saved[$number] = @{
            REQ_TIT='old-title'; REQ_RMK="old request <p>&amp; 50% + 한글`r`n둘째"; GUBUN='PJT'; STS='1'; PGM_MD=''; ACT_GUBUN='';
            PGM_NM='old-program'; PGM_ID='old-id'; ACT_DT_FR='2026-10-01'; ACT_DT='2026-10-22'; ACT_TIME='30'; ACT_UID='worker';
            ACT_RMK="<p>웹 전용 &amp; 100% + 처리</p>`r`n둘째"; ADM_REMARK='web-report-title'; ACCPT_RMK='web-acceptance'; unique_key='original-upload-key'
        }
        $existingViews[$number] = 0
    }
    $saved['RQ9199']['GUBUN'] = 'MA'
    try {
        Write-Output 'READY'
        while ($true) {
            $client = $listener.AcceptTcpClient()
            $client.ReceiveTimeout = 10000
            $stream = $client.GetStream()
            $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::ASCII, $false, 1024, $true)
            $requestLine = $reader.ReadLine()
            if (-not $requestLine) { $client.Close(); continue }
            $requestParts = $requestLine.Split(' ')
            $method = $requestParts[0]
            $target = $requestParts[1]
            Add-Content -LiteralPath $serverTrace -Value ($method + ' ' + $target)
            $headers = @{}
            while (($line = $reader.ReadLine()) -ne '') {
                if ($null -eq $line) { break }
                $pair = $line.Split([char[]]@(':'), 2)
                $headers[$pair[0].ToLowerInvariant()] = $pair[1].Trim()
            }
            $body = ''
            if ($headers.ContainsKey('content-length')) {
                $remaining = [int]$headers['content-length']
                $buffer = New-Object char[] $remaining
                $offset = 0
                while ($offset -lt $remaining) {
                    $count = $reader.Read($buffer, $offset, $remaining - $offset)
                    if ($count -eq 0) { break }
                    $offset += $count
                }
                $body = [string]::new($buffer, 0, $offset)
            }
            $postedValues = @{}
            foreach ($item in $body.Split('&')) {
                $pair = $item.Split([char[]]@('='), 2)
                if ($pair.Count -eq 2) { $postedValues[[Net.WebUtility]::UrlDecode($pair[0])] = [Net.WebUtility]::UrlDecode($pair[1]) }
            }
            $status = '200 OK'
            $extra = ''
            $html = '<html><body>fixture</body></html>'
            $stopFixture = $false
            $cookie = [string]$headers['cookie']
            if ($target -eq '/__stop') {
                $stopFixture = $true
            } elseif ($target -eq '/support/') {
                $status = '302 Found'
                $extra = "Location: /support/login.aspx`r`nSet-Cookie: fixture=one; Path=/support/`r`n"
                $html = ''
            } elseif ($target -eq '/support/login.aspx' -and $method -eq 'GET') {
                $html = '<form method="post" action="./login.aspx"><input type="hidden" name="__VIEWSTATE" value="login+state&amp;1"><input name="Login1$UserName"><input type="password" name="Login1$Password"><input type="submit" name="Login1$LoginButton" value="Login"></form>'
            } elseif ($target -eq '/support/login.aspx' -and $method -eq 'POST') {
                if ($cookie -notmatch 'fixture=one' -or $postedValues['__VIEWSTATE'] -ne 'login+state&1' -or $postedValues['Login1$UserName'] -ne 'fixture-user' -or $postedValues['Login1$Password'] -ne 'fixture-password') { throw 'fixture login mismatch' }
                $status = '302 Found'
                $extra = "Location: /support/Default.aspx`r`nSet-Cookie: auth=yes; Path=/support/; HttpOnly`r`n"
                $html = ''
            } elseif ($target -eq '/support/Default.aspx') {
                if ($cookie -notmatch 'auth=yes') { throw 'fixture auth cookie absent' }
            } elseif ($target -eq '/support/C0/c0210_in.aspx' -and $method -eq 'GET') {
                if ($cookie -notmatch 'auth=yes') { throw 'fixture auth cookie absent on registration' }
                $sequence++
                $html = Registration ('RQ' + (9000 + $sequence)) ('view' + $sequence + '+a&b')
            } elseif ($target -eq '/support/C0/c0210_in.aspx' -and $method -eq 'POST') {
                $savePosts++
                if ($postedValues['__VIEWSTATE'] -ne ('view' + $sequence + '+a&b') -or $postedValues['btnSave.x'] -ne '1' -or $postedValues['btnSave.y'] -ne '1') { throw 'fresh-state or image button mismatch' }
                if (-not $postedValues.ContainsKey('ACT_TIME') -or $postedValues['ACT_TIME'] -ne '') { throw 'blank ACT_TIME was omitted or defaulted' }
                $saved['RQ' + (9000 + $sequence)] = $postedValues
                if ($postedValues['REQ_TIT'] -eq 'fixture-connection-drop') {
                    $client.Close()
                    Write-Output ('SAVE_POSTS=' + $savePosts)
                    continue
                }
                $status = '302 Found'
                $extra = "Location: /support/C0/c0200_02.aspx?REQ_NO=RQ" + (9000 + $sequence) + "`r`n"
                $html = ''
            } elseif ($target -match '^/support/C0/c0200_02\.aspx\?REQ_NO=([^&]+)' -and $method -eq 'POST') {
                $number = [Net.WebUtility]::UrlDecode($Matches[1])
                if (-not $existingViews.ContainsKey($number)) { throw 'update must target a seeded existing record' }
                $updatePosts++
                $record = $saved[$number]
                if ($postedValues['__VIEWSTATE'] -ne ('existing-' + $number + '-' + $existingViews[$number]) -or $postedValues['seqNum'] -ne $number) { throw 'existing fresh state / request ID mismatch' }
                if ($postedValues.ContainsKey('REQ_UID') -or $postedValues.ContainsKey('btnSave.x')) { throw 'existing save cannot write requester or use creation image button' }
                foreach ($name in @('ADM_REMARK','ACCPT_RMK','unique_key')) {
                    if ($postedValues[$name] -cne $record[$name]) { throw ('web-only value changed: ' + $name) }
                }
                if ([Net.WebUtility]::UrlDecode($postedValues['ACT_RMK']) -cne $record['ACT_RMK']) { throw 'web-only ACT_RMK was changed or incorrectly encoded' }
                $eventTarget = $postedValues['__EVENTTARGET']
                if ($number -eq 'RQ9102' -and $eventTarget -eq 'btnSave') {
                    $client.Close()
                    Add-Content -LiteralPath $serverTrace -Value 'intentional dropped processing save before persistence'
                    continue
                }
                if ($eventTarget -eq 'btnSaveReq') {
                    $record['REQ_TIT'] = $postedValues['REQ_TIT']
                    $record['REQ_RMK'] = [Net.WebUtility]::UrlDecode($postedValues['REQ_RMK'])
                } elseif ($eventTarget -eq 'btnSave') {
                    if ([Net.WebUtility]::UrlDecode($postedValues['REQ_RMK']) -cne $record['REQ_RMK']) { throw 'processing save changed request body' }
                    foreach ($name in @('GUBUN','STS','PGM_MD','ACT_GUBUN','PGM_NM','PGM_ID','ACT_DT_FR','ACT_DT','ACT_TIME','ACT_UID')) { $record[$name] = $postedValues[$name] }
                } else { throw 'unexpected update event target' }
                if ($number -eq 'RQ9103') {
                    $client.Close()
                    Write-Output ('UPDATE_POSTS=' + $updatePosts)
                    continue
                }
                $html = '<html>saved</html>'
            } elseif ($target -match '^/support/C0/c0200_02\.aspx\?REQ_NO=([^&]+)' -and $method -eq 'GET') {
                $number = [Net.WebUtility]::UrlDecode($Matches[1])
                if ($saved.ContainsKey($number)) {
                    $record = $saved[$number]
                    if ($existingViews.ContainsKey($number)) {
                        $existingViews[$number]++
                        $record['unique_key'] = 'upload-state-' + $existingViews[$number]
                        if ($number -eq 'RQ9101' -and $existingViews[$number] -eq 2) {
                            $record['ADM_REMARK'] = 'web-report-updated-after-preview'
                            $record['ACT_RMK'] += "`r`n웹에서만 추가한 내용 % + &"
                        }
                        if ($number -eq 'RQ9104' -and $existingViews[$number] -eq 2) { $record['REQ_TIT'] = 'remote-title' }
                        $html = Existing $number $record ('existing-' + $number + '-' + $existingViews[$number])
                    } else {
                        $html = '<form><span id="REQ_NO">' + $number + '</span><span id="REQ_UID">u1</span><textarea name="REQ_RMK" id="REQ_RMK">' + "`r`n" + [Net.WebUtility]::HtmlEncode($record['REQ_RMK']) + '</textarea>'
                        foreach ($name in @('REQ_TIT','GUBUN','STS','PGM_NM','PGM_ID','ACT_TIME')) { $html += Field $name $record[$name] }
                        $html += '</form>'
                    }
                } else { $html = '<form>not found</form>' }
            } else { throw ('unexpected fixture request ' + $method + ' ' + $target) }
            $payload = [Text.Encoding]::UTF8.GetBytes($html)
            $responseHeader = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $status`r`n$extra" + "Content-Type: text/html; charset=utf-8`r`nContent-Length: " + $payload.Length + "`r`nConnection: close`r`n`r`n")
            $stream.Write($responseHeader, 0, $responseHeader.Length)
            $stream.Write($payload, 0, $payload.Length)
            $stream.Flush()
            Add-Content -LiteralPath $serverTrace -Value ('response ' + $status)
            $reader.Dispose()
            $client.Close()
            if ($stopFixture) { break }
        }
    } catch {
        Add-Content -LiteralPath $serverTrace -Value ('ERROR: ' + $_.Exception.Message)
        throw
    } finally { $listener.Stop() }
}
$taskExcel = $null
$taskBook = $null
$testFailure = $null
$tracePath = Join-Path ([IO.Path]::GetTempPath()) ('web-transport-' + [guid]::NewGuid().ToString('N') + '.trace')
try {
    $ready = $false
    for ($i = 0; $i -lt 40; $i++) {
        if ($fixtureJob.ChildJobs[0].Output.Count -gt 0) { $ready = $true; break }
        if ($fixtureJob.State -eq 'Failed') { Receive-Job $fixtureJob; throw 'fixture failed to start' }
        Start-Sleep -Milliseconds 100
    }
    if (-not $ready) { throw 'fixture startup timed out' }
    Write-Output 'Web transport regression: starting isolated Excel application'
    $taskExcel = New-Object -ComObject Excel.Application
    $taskExcel.Visible = $false
    $taskExcel.DisplayAlerts = $false
    $taskExcel.EnableEvents = $false
    $taskBook = $taskExcel.Workbooks.Add()
    Write-Output 'Web transport regression: importing modules into an unsaved blank workbook'
    foreach ($relative in @('vba-files/Module/modWebTransport.bas', 'tests/WebTransportRegression.bas')) {
        $source = [IO.File]::ReadAllText((Join-Path $taskRoot $relative), [Text.Encoding]::UTF8)
        $name = [regex]::Match($source, '(?m)^Attribute VB_Name = "([^"]+)"').Groups[1].Value
        $source = $source.TrimStart([char]0xFEFF) -replace '(?m)^Attribute VB_Name = .*\r?\n', ''
        $component = $taskBook.VBProject.VBComponents.Add(1)
        $component.Name = $name
        $component.CodeModule.AddFromString($source)
    }
    Write-Output 'Web transport regression: running VBA and loopback fixture checks'
    $fixtureBase = if ($NoNetwork) { '' } else { 'http://127.0.0.1:' + $fixturePort + '/support/' }
    $result = $taskExcel.Run("'" + $taskBook.Name + "'!RunWebTransportRegression", $fixtureBase, $tracePath)
    Write-Output $result
    if ([string]$result -notlike 'PASS:*') { throw [string]$result }
    $fixtureOutput = @($fixtureJob.ChildJobs[0].Output | ForEach-Object { [string]$_ })
    if (-not $NoNetwork) {
        if ($fixtureOutput -notcontains 'SAVE_POSTS=2') { throw 'fixture did not observe exactly two POSTs (success + unknown), or a stale option was submitted' }
        if ($fixtureOutput -notcontains 'UPDATE_POSTS=5') { throw 'fixture did not observe exactly five update POSTs (two-tab success + partial + confirmed lost response), or a rejected update was submitted' }
        Write-Output 'PASS: exactly two registration / five update POSTs; web-only values preserved; no automatic POST retry; no live website save'
    }
} catch {
    $testFailure = $_
} finally {
    if ($testFailure) {
        if (Test-Path -LiteralPath $tracePath) { Get-Content -LiteralPath $tracePath }
        if (Test-Path -LiteralPath $serverTrace) { Get-Content -LiteralPath $serverTrace }
    }
    if ($taskBook) { try { $taskBook.Close($false) } catch {} }
    if ($taskExcel) { try { $taskExcel.Quit() } catch {} }
    foreach ($comObject in @($taskBook, $taskExcel)) { if ($comObject) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } }
    if ($fixtureJob) {
        if ($fixtureJob.State -eq 'Running') {
            try {
                $stopClient = [Net.Sockets.TcpClient]::new()
                $stopClient.ReceiveTimeout = 2000
                $stopClient.Connect([Net.IPAddress]::Loopback, $fixturePort)
                $stopBytes = [Text.Encoding]::ASCII.GetBytes("GET /__stop HTTP/1.1`r`nHost: localhost`r`nConnection: close`r`n`r`n")
                $stopStream = $stopClient.GetStream()
                $stopStream.Write($stopBytes, 0, $stopBytes.Length)
                $stopBuffer = New-Object byte[] 64
                [void]$stopStream.Read($stopBuffer, 0, $stopBuffer.Length)
                $stopClient.Close()
            } catch {}
            [void](Wait-Job $fixtureJob -Timeout 3)
        }
        if ($fixtureJob.State -eq 'Failed') { Receive-Job $fixtureJob -ErrorAction Continue }
        if ($fixtureJob.State -eq 'Running') { Stop-Job $fixtureJob }
        Remove-Job $fixtureJob -Force
    }
    if (Test-Path -LiteralPath $tracePath) { Remove-Item -LiteralPath $tracePath -Force }
    if (Test-Path -LiteralPath $serverTrace) { Remove-Item -LiteralPath $serverTrace -Force }
}
if ($testFailure) { throw $testFailure }
