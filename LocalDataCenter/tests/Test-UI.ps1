param([string]$Browser = 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', [switch]$StatusOnly, [switch]$OCROnly, [string]$PreviewPdf)
$ErrorActionPreference = 'Stop'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ldc-ui-' + [guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($testRoot) | Out-Null
$server = $null; $browserProcess = $null; $socket = $null
try {
    $script = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../LocalDataCenter.ps1'))
    $output = Join-Path $testRoot 'server.log'
    $server = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -DataRoot `"$testRoot/data`" -NoBrowser" -WindowStyle Hidden -RedirectStandardOutput $output -RedirectStandardError (Join-Path $testRoot 'server-error.log') -PassThru
    $origin = $null
    for ($i=0; $i -lt 100; $i++) {
        Start-Sleep -Milliseconds 100
        if (Test-Path $output) { $log = Get-Content -Raw $output; if ($log -match 'http://127\.0\.0\.1:\d+') { $origin = $Matches[0]; break } }
        if ($server.HasExited) { throw ([IO.File]::ReadAllText((Join-Path $testRoot 'server-error.log'))) }
    }
    if (!$origin) { throw 'Server did not start' }
    $profile = Join-Path $testRoot 'browser'
    $browserProcess = Start-Process $Browser -ArgumentList "--headless=new --disable-gpu --no-first-run --no-default-browser-check --remote-debugging-port=0 --remote-allow-origins=* --user-data-dir=`"$profile`" about:blank" -WindowStyle Hidden -PassThru
    $portFile = Join-Path $profile 'DevToolsActivePort'
    for ($i=0; $i -lt 100 -and !(Test-Path $portFile); $i++) { Start-Sleep -Milliseconds 100 }
    $port = ([IO.File]::ReadAllLines($portFile))[0]
    $target = $null
    for ($i=0; $i -lt 50; $i++) {
        $targets = Invoke-RestMethod "http://127.0.0.1:$port/json/list"
        $target = $targets | Where-Object type -eq 'page' | Select-Object -First 1
        if ($target.webSocketDebuggerUrl) { break }
        Start-Sleep -Milliseconds 100
    }
    if (!$target) { throw 'Browser did not create a page target' }
    $socket = [Net.WebSockets.ClientWebSocket]::new()
    $null = $socket.ConnectAsync([uri]$target.webSocketDebuggerUrl,[Threading.CancellationToken]::None).GetAwaiter().GetResult()
    $script:commandId = 0
    function Cdp($method, $parameters = @{}) {
        $script:commandId++
        $requestId = $script:commandId
        $json = @{id=$requestId;method=$method;params=$parameters} | ConvertTo-Json -Depth 25 -Compress
        $bytes = [Text.Encoding]::UTF8.GetBytes($json)
        $socket.SendAsync([ArraySegment[byte]]::new($bytes),[Net.WebSockets.WebSocketMessageType]::Text,$true,[Threading.CancellationToken]::None).GetAwaiter().GetResult()
        do {
            $message = [IO.MemoryStream]::new()
            do {
                $buffer = New-Object byte[] 65536
                $timeout = [Threading.CancellationTokenSource]::new(60000)
                try { $received = $socket.ReceiveAsync([ArraySegment[byte]]::new($buffer),$timeout.Token).GetAwaiter().GetResult() } finally { $timeout.Dispose() }
                $message.Write($buffer,0,$received.Count)
            } until ($received.EndOfMessage)
            $reply = [Text.Encoding]::UTF8.GetString($message.ToArray()) | ConvertFrom-Json
            if ($reply.method -eq 'Runtime.exceptionThrown') { [Console]::Error.WriteLine(($reply.params | ConvertTo-Json -Depth 10 -Compress)) }
            if ($reply.method -eq 'Runtime.consoleAPICalled' -and $reply.params.args[0].value -like 'TEST:*') { Write-Output $reply.params.args[0].value | Out-Host }
            if ($reply.method -eq 'Log.entryAdded' -and $reply.params.entry.level -eq 'error') { [Console]::Error.WriteLine($reply.params.entry.text) }
            $message.Dispose()
        } until ($reply.id -eq $requestId)
        if ($reply.error) { throw ($reply.error | ConvertTo-Json -Compress) }
        return $reply.result
    }
    $null = Cdp 'Page.enable'
    $null = Cdp 'Runtime.enable'
    $null = Cdp 'Log.enable'
    $compiled = Cdp 'Runtime.compileScript' @{expression=[IO.File]::ReadAllText((Join-Path $PSScriptRoot '../web/app.js'));sourceURL='app.js';persistScript=$false}
    if ($compiled.exceptionDetails) { throw ($compiled.exceptionDetails | ConvertTo-Json -Depth 10) }
    $compiled = Cdp 'Runtime.compileScript' @{expression=[IO.File]::ReadAllText((Join-Path $PSScriptRoot '../web/features.js'));sourceURL='features.js';persistScript=$false}
    if ($compiled.exceptionDetails) { throw ($compiled.exceptionDetails | ConvertTo-Json -Depth 10) }
    $null = Cdp 'Emulation.setDeviceMetricsOverride' @{width=1440;height=1050;deviceScaleFactor=1;mobile=$false}
    $null = Cdp 'Page.navigate' @{url=$origin}
    Start-Sleep -Seconds 2
    if ($PreviewPdf) {
        $fixtureId = [guid]::NewGuid().ToString('N') + '.bin'
        Copy-Item -LiteralPath $PreviewPdf -Destination (Join-Path $testRoot "data/Attachments/$fixtureId")
        $previewCheck = '(async()=>{for(let i=0;i<500 && (typeof data==="undefined" || !data.capabilities);i++)await new Promise(r=>setTimeout(r,100));await openPreview({id:"' + $fixtureId + '",name:"restored.pdf"});for(let i=0;i<200 && !document.querySelector(".pdf-page canvas") && !document.querySelector("#preview-status").textContent.includes("失敗") && !document.querySelector("#preview-status").textContent.includes("無法預覽");i++)await new Promise(r=>setTimeout(r,100));return {status:document.querySelector("#preview-status").textContent,pages:previewSession?.pdf?.numPages,canvases:document.querySelectorAll(".pdf-page canvas").length};})()'
        $result = Cdp 'Runtime.evaluate' @{expression=$previewCheck;awaitPromise=$true;returnByValue=$true}
        if ($result.exceptionDetails) { throw ($result.exceptionDetails | ConvertTo-Json -Depth 10) }
        $result.result.value | ConvertTo-Json -Depth 8
        return
    }
    if ($OCROnly) {
        $setup = '(async()=>{for(let i=0;i<500 && (typeof data==="undefined" || !data.capabilities);i++) await new Promise(r=>setTimeout(r,100));projectId=(await save({kind:"project",name:"OCR tests"})).id;sectionId=(await save({kind:"section",projectId,name:"Tests"})).id;eventId=(await save({kind:"event",sectionId,title:"OCR target",date:"2026-11-10",html:"",files:[],schedules:[]})).id;await refresh();})()'
        $result = Cdp 'Runtime.evaluate' @{expression=$setup;awaitPromise=$true;returnByValue=$true}
        if ($result.exceptionDetails) { throw ($result.exceptionDetails | ConvertTo-Json -Depth 10) }
        $result = Cdp 'Runtime.evaluate' @{expression=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ocr-checks.js'));awaitPromise=$true;returnByValue=$true}
        if ($result.exceptionDetails) { throw ($result.exceptionDetails | ConvertTo-Json -Depth 10) }
        $result.result.value | ConvertTo-Json -Depth 8
        return
    }
    if ($StatusOnly) { $null = Cdp 'Runtime.evaluate' @{expression='window.statusOnly = true'} }
    $checks = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'ui-checks.js'))
    $result = Cdp 'Runtime.evaluate' @{expression=$checks;awaitPromise=$true;returnByValue=$true}
    if ($result.exceptionDetails) { throw ($result.exceptionDetails | ConvertTo-Json -Depth 10) }
    $result.result.value | ConvertTo-Json -Depth 8
    $result = Cdp 'Runtime.evaluate' @{expression=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'calendar-checks.js'));awaitPromise=$true;returnByValue=$true}
    if ($result.exceptionDetails) { throw ($result.exceptionDetails | ConvertTo-Json -Depth 10) }
    $result.result.value | ConvertTo-Json -Depth 8
    if (!$StatusOnly) {
        foreach ($checkFile in @('feature-checks.js','ocr-checks.js')) {
            $checks = [IO.File]::ReadAllText((Join-Path $PSScriptRoot $checkFile))
            $result = Cdp 'Runtime.evaluate' @{expression=$checks;awaitPromise=$true;returnByValue=$true}
            if ($result.exceptionDetails) { throw ($result.exceptionDetails | ConvertTo-Json -Depth 10) }
            $result.result.value | ConvertTo-Json -Depth 8
        }
        . (Join-Path $PSScriptRoot 'table-mouse-checks.ps1')
    }
    foreach ($width in @(1440,900,390)) {
        $null = Cdp 'Emulation.setDeviceMetricsOverride' @{width=$width;height=1050;deviceScaleFactor=1;mobile=$false}
        $layout = Cdp 'Runtime.evaluate' @{expression='({width:innerWidth,scroll:document.documentElement.scrollWidth,calendar:document.querySelector("#calendar").getBoundingClientRect().width})';returnByValue=$true}
        if ($layout.result.value.scroll -gt $width) { throw "Horizontal overflow at width $width" }
        $shot = Cdp 'Page.captureScreenshot' @{format='png'}
        [IO.File]::WriteAllBytes((Join-Path $testRoot "layout-$width.png"),[Convert]::FromBase64String($shot.data))
        $result = Cdp 'Runtime.evaluate' @{expression='openScheduleCalendar(); openCalendarDay(localDate(new Date())); (()=>{const d=$("schedule-calendar"),r=d.getBoundingClientRect();return {width:r.width,height:r.height,scroll:d.scrollWidth,client:d.clientWidth,right:r.right,left:r.left};})()';returnByValue=$true}
        $geometry = $result.result.value
        if ($result.exceptionDetails -or $geometry.scroll -gt ($geometry.client + 1) -or $geometry.right -gt $width -or $geometry.left -lt 0) { throw "Calendar overflow at width $width" }
        if ($width -eq 1440 -and [Math]::Abs($geometry.width - ($width * 0.8)) -gt 2) { throw 'Calendar modal should occupy 80% of desktop width' }
        $shot = Cdp 'Page.captureScreenshot' @{format='png'}
        [IO.File]::WriteAllBytes((Join-Path $testRoot "calendar-$width.png"),[Convert]::FromBase64String($shot.data))
        $null = Cdp 'Input.dispatchKeyEvent' @{type='keyDown';key='Escape';code='Escape';windowsVirtualKeyCode=27}
        $null = Cdp 'Input.dispatchKeyEvent' @{type='keyUp';key='Escape';code='Escape';windowsVirtualKeyCode=27}
        $result = Cdp 'Runtime.evaluate' @{expression='$("schedule-calendar").open';returnByValue=$true}
        if ($result.result.value) { throw 'Escape did not dismiss calendar modal' }
        $editorCheck = Cdp 'Runtime.evaluate' @{expression='openEditor(data.Events.find(e=>e.id===eventId)); ({width:document.querySelector("#edit").clientWidth,scroll:document.querySelector("#edit").scrollWidth})';returnByValue=$true}
        if ($editorCheck.result.value.scroll -gt $editorCheck.result.value.width) { throw "Editor overflow at width $width" }
        $null = Cdp 'Runtime.evaluate' @{expression='document.querySelector("#edit").close()'}
    }
    Write-Output "PASS: desktop, tablet and mobile have no horizontal overflow. Screenshots: $testRoot"
    foreach ($expected in @(3,0)) {
        Stop-Process -Id $server.Id -Force
        $server.WaitForExit()
        $output = Join-Path $testRoot "restart-$expected.log"
        $server = Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$script`" -DataRoot `"$testRoot/data`" -NoBrowser" -WindowStyle Hidden -RedirectStandardOutput $output -RedirectStandardError (Join-Path $testRoot "restart-$expected-error.log") -PassThru
        $origin = $null
        for ($i=0; $i -lt 100; $i++) {
            Start-Sleep -Milliseconds 100
            if (Test-Path $output) { $log = Get-Content -Raw $output; if ($log -match 'http://127\.0\.0\.1:\d+') { $origin = $Matches[0]; break } }
            if ($server.HasExited) { throw 'Restarted server exited unexpectedly' }
        }
        if (!$origin) { throw 'Server restart timed out' }
        $null = Cdp 'Page.navigate' @{url=$origin}
        Start-Sleep -Seconds 2
        $expression = '(async()=>{ for(let i=0;i<500 && (document.readyState!=="complete" || typeof data==="undefined");i++) await new Promise(r=>setTimeout(r,100)); const all=await api("/api/data"); const event=all.Events.find(e=>e.title==="Multiple schedules"); const count=event.schedules.length; const schedule=event.schedules[0]; if(count===3) await save({...event,kind:"event",schedules:[]}); return {count,start:event.start,end:event.end,status:event.status,schedule,deletedPresent:all.Events.some(e=>e.title==="Delete fixture")}; })()'
        $restart = Cdp 'Runtime.evaluate' @{expression=$expression;awaitPromise=$true;returnByValue=$true}
        if ($restart.exceptionDetails -or $restart.result.value.count -ne $expected) { throw "Restart did not preserve $expected schedules" }
        if ($expected -eq 0 -and ($restart.result.value.start -or $restart.result.value.end)) { throw 'Cleared schedules retained legacy dates' }
        if ($restart.result.value.status -ne 'completed') { throw 'Restart lost completed status' }
        if ($restart.result.value.deletedPresent) { throw 'Restart restored deleted event' }
        if ($expected -eq 3 -and (!$restart.result.value.schedule.completed -or $restart.result.value.schedule.completedDate -ne '2026-11-08' -or $restart.result.value.schedule.files.Count -ne 2)) { throw 'Restart lost schedule details' }
    }
    Write-Output 'PASS: full server restarts preserve multiple schedules and explicitly cleared schedules.'
} finally {
    if ($socket) { $socket.Dispose() }
    if ($browserProcess -and !$browserProcess.HasExited) { Stop-Process -Id $browserProcess.Id -Force }
    if ($server -and !$server.HasExited) { Stop-Process -Id $server.Id -Force }
}
