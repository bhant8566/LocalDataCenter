# Local Data Center | Windows PowerShell 5.1 + Microsoft Edge / Chrome.
# 啟動：在 PowerShell 進入腳本所在資料夾，再執行 .\LocalDataCenter.ps1
# 預設資料位置：桌面\LocalDataCenter。保留 Excel 和 Attachments 即可移轉資料。
# 停止：在 PowerShell 視窗按 Ctrl+C。
[CmdletBinding()]
param(
    [string]$DataRoot = (Join-Path ([Environment]::GetFolderPath('DesktopDirectory')) 'LocalDataCenter'),
    [switch]$NoBrowser
)
$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$origin = 'http://127.0.0.1'
$token = [Guid]::NewGuid().ToString('N')
$workbook = Join-Path $DataRoot 'LocalDataCenter.xlsx'
$attachments = Join-Path $DataRoot 'Attachments'

$schema = [ordered]@{
    Projects = @('id','name')
    Sections = @('id','projectId','name')
    Events = @('id','sectionId','title','date','start','end')
    Content = @('eventId','part','data')
}

function Read-Data {
    $data = @{ Projects = @(); Sections = @(); Events = @() }
    if (!(Test-Path -LiteralPath $workbook)) { return $data }
    $package = Open-ExcelPackage -Path $workbook
    try {
        $content = @{}
        foreach ($name in $schema.Keys) {
            $sheet = $package.Workbook.Worksheets[$name]
            if (!$sheet -or !$sheet.Dimension) { throw "Excel 格式不符。找不到 $name 工作表。" }
            $columns = $schema[$name]
            for ($i = 0; $i -lt $columns.Count; $i++) {
                if ($sheet.Cells[1, ($i + 1)].Text -ne $columns[$i]) {
                    throw "Excel 格式不符。請確認 $name 的表頭。"
                }
            }
            $rows = [Collections.Generic.List[object]]::new()
            for ($r = 2; $r -le $sheet.Dimension.End.Row; $r++) {
                if (!$sheet.Cells[$r,1].Text) { continue }
                $row = [ordered]@{}
                for ($i = 0; $i -lt $columns.Count; $i++) {
                    $row[$columns[$i]] = $sheet.Cells[$r, ($i + 1)].Text
                }
                $rows.Add([pscustomobject]$row)
            }
            if ($name -eq 'Content') {
                foreach ($group in ($rows | Group-Object eventId)) {
                    $encoded = ($group.Group | Sort-Object { [int]$_.part } | ForEach-Object { $_.data }) -join ''
                    $content[$group.Name] = $utf8.GetString([Convert]::FromBase64String($encoded)) | ConvertFrom-Json
                }
            } else { $data[$name] = @($rows.ToArray()) }
        }
        foreach ($event in $data.Events) {
            $body = $content[$event.id]
            $event | Add-Member NoteProperty html ([string]$body.html)
            $event | Add-Member NoteProperty files @($body.files | Where-Object { $_ })
            $schedules = @()
            if ($body -and $body.PSObject.Properties['schedules']) { $schedules = @($body.schedules) }
            elseif ($event.start -and $event.end) {
                $schedules = @([pscustomobject]@{name='原有時程'; start=$event.start; end=$event.end})
            }
            for ($s = 0; $s -lt $schedules.Count; $s++) {
                $schedule = $schedules[$s]
                if (!$schedule.id) { $schedule | Add-Member NoteProperty id ("{0}-{1}" -f $event.id,$s) -Force }
                if (!$schedule.PSObject.Properties['notes']) { $schedule | Add-Member NoteProperty notes '' }
                if (!$schedule.PSObject.Properties['files']) { $schedule | Add-Member NoteProperty files @() }
                if (!$schedule.PSObject.Properties['completed']) { $schedule | Add-Member NoteProperty completed $false }
                if (!$schedule.PSObject.Properties['completedDate']) { $schedule | Add-Member NoteProperty completedDate '' }
            }
            $event | Add-Member NoteProperty schedules $schedules
            $status = 'active'
            if ($body.status -eq 'completed') { $status = 'completed' }
            $event | Add-Member NoteProperty status $status
        }
        return $data
    } finally { $package.Dispose() }
}

function Save-Data($data) {
    $temp = Join-Path $DataRoot ("pending-{0}.xlsx" -f [Guid]::NewGuid())
    $package = Open-ExcelPackage -Path $temp -Create
    try {
        $parts = [Collections.Generic.List[object]]::new()
        foreach ($event in $data.Events) {
            $json = ConvertTo-Json @{html=$event.html; files=@($event.files); schedules=@($event.schedules); status=$event.status} -Depth 8 -Compress
            $encoded = [Convert]::ToBase64String($utf8.GetBytes($json))
            for ($offset = 0; $offset -lt $encoded.Length; $offset += 30000) {
                $parts.Add(@{eventId=$event.id; part=[int]($offset / 30000); data=$encoded.Substring($offset,[Math]::Min(30000,$encoded.Length-$offset))})
            }
        }
        foreach ($name in $schema.Keys) {
            $sheet = $package.Workbook.Worksheets.Add($name)
            $columns = $schema[$name]
            for ($i = 0; $i -lt $columns.Count; $i++) { $sheet.Cells[1, ($i + 1)].Value = $columns[$i] }
            $sheet.Row(1).Style.Font.Bold = $true
            $r = 2
            $rows = $data[$name]
            if ($name -eq 'Content') { $rows = $parts }
            foreach ($row in $rows) {
                for ($i = 0; $i -lt $columns.Count; $i++) {
                    $column = $columns[$i]
                    $sheet.Cells[$r, ($i + 1)].Value = [string]$row.$column
                }
                $r++
            }
        }
        $package.Save()
        $package.Dispose()
        if ([IO.File]::Exists($workbook)) { [IO.File]::Replace($temp,$workbook,(Join-Path $DataRoot "LocalDataCenter.previous.xlsx")) }
        else { [IO.File]::Move($temp,$workbook) }
    } finally {
        $package.Dispose()
        if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
    }
}

function Read-Request($stream) {
    $header = [IO.MemoryStream]::new()
    $tail = 0
    while ($header.Length -lt 16384) {
        $byte = $stream.ReadByte()
        if ($byte -lt 0) { return }
        $header.WriteByte($byte)
        $tail = (($tail -shl 8) -bor $byte) -band 0xffffffffL
        if ($tail -eq 0x0d0a0d0a) { break }
    }
    if ($tail -ne 0x0d0a0d0a) { throw 'HTTP header 過長' }
    $lines = [Text.Encoding]::ASCII.GetString($header.ToArray()).Split("`n")
    $request = $lines[0].Trim().Split(' ')
    if ($request.Count -ne 3 -or $request[1] -notmatch '^/[^?# ]*$') { throw 'HTTP request 格式錯誤' }
    $headers = @{}
    foreach ($line in ($lines | Select-Object -Skip 1)) {
        if (!$line.Trim()) { continue }
        $pair = $line.Trim().Split(':',2)
        if ($pair.Count -ne 2 -or $headers.ContainsKey($pair[0])) { throw 'HTTP header 格式錯誤' }
        $headers[$pair[0]] = $pair[1].Trim()
    }
    return @{ Method=$request[0]; Path=$request[1]; Headers=$headers }
}

function Read-Body($stream, $request, [int]$limit) {
    $length = 0
    if ($request.Headers['Transfer-Encoding'] -or ![int]::TryParse($request.Headers['Content-Length'],[ref]$length) -or $length -lt 0 -or $length -gt $limit) {
        throw '檔案或表單超過允許大小，或請求格式錯誤。'
    }
    $bytes = New-Object byte[] $length
    $position = 0
    while ($position -lt $length) {
        $read = $stream.Read($bytes,$position,$length-$position)
        if ($read -eq 0) { throw '上傳中斷。' }
        $position += $read
    }
    return ,$bytes
}

function Send-Response($stream, $status, $type, [byte[]]$bytes, $extra = '', $cacheControl = 'no-store') {
    $script:reuseClient = $status -eq '200 OK'
    $connection = 'close'
    if ($script:reuseClient) { $connection = 'keep-alive' }
    $headers = "HTTP/1.1 $status`r`nContent-Type: $type`r`nContent-Length: $($bytes.Length)`r`nConnection: $connection`r`nX-Content-Type-Options: nosniff`r`nReferrer-Policy: no-referrer`r`nCache-Control: $cacheControl`r`n$extra`r`n"
    $raw = [Text.Encoding]::ASCII.GetBytes($headers)
    $stream.Write($raw,0,$raw.Length)
    $stream.Write($bytes,0,$bytes.Length)
    $stream.Flush()
}

function Send-Json($stream, $value, $status = '200 OK') {
    Send-Response $stream $status 'application/json; charset=utf-8' ($utf8.GetBytes((ConvertTo-Json -InputObject $value -Depth 12 -Compress)))
}

function Check-Date([string]$value) {
    $parsed = [datetime]::MinValue
    return [datetime]::TryParseExact($value,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$parsed)
}

function Validate-Files($files) {
    if (@($files).Count -gt 100) { throw '每個項目最多 100 個檔案。' }
    foreach ($file in $files) {
        if ($file.id -notmatch '^[a-f0-9]{32}\.(png|jpg|gif|webp|bin)$' -or !$file.name -or $file.name.Length -gt 255 -or !(Test-Path -LiteralPath (Join-Path $attachments $file.id) -PathType Leaf)) {
            throw '附件不存在或格式錯誤。'
        }
    }
}

function Convert-Schedule($schedule) {
    $name = ([string]$schedule.name).Trim()
    if (!$name -or $name.Length -gt 80) { throw '時程名稱須為 1 至 80 個字。' }
    if (!(Check-Date $schedule.start) -or !(Check-Date $schedule.end) -or $schedule.end -lt $schedule.start) {
        throw '請填寫有效的時程日期，且結束日期不可早於開始日期。'
    }
    $id = [string]$schedule.id
    if (!$id) { $id = [Guid]::NewGuid().ToString('N') }
    if ($id -notmatch '^[a-f0-9]{32}(-[0-9]+)?$') { throw '時程識別碼無效。' }
    $notes = [string]$schedule.notes
    if ($notes.Length -gt 500000) { throw '時程文字超過 50 萬字元。' }
    $files = @($schedule.files | Where-Object { $_ })
    Validate-Files $files
    $completed = $false
    if ($schedule.PSObject.Properties['completed']) {
        if ($schedule.completed -isnot [bool]) { throw '完成狀態必須是布林值。' }
        $completed = $schedule.completed
    }
    $completedDate = ''
    if ($completed) {
        if (!(Check-Date $schedule.completedDate)) { throw '請先手動選擇有效的完成日期。' }
        $completedDate = [string]$schedule.completedDate
    }
    return [pscustomobject]@{id=$id; name=$name; start=[string]$schedule.start; end=[string]$schedule.end; notes=$notes; files=$files; completed=$completed; completedDate=$completedDate}
}

function Update-Data($payload) {
    $data = Read-Data
    $beforeFiles = @(Get-AttachmentIds $data)
    $id = [string]$payload.id
    switch ($payload.kind) {
        { $_ -in 'project','section' } {
            $name = ([string]$payload.name).Trim()
            if (!$name -or $name.Length -gt 80) { throw '名稱須為 1 至 80 個字。' }
            $collection = 'Projects'
            if ($payload.kind -eq 'section') { $collection = 'Sections' }
            if ($id) {
                $row = $data[$collection] | Where-Object id -eq $id | Select-Object -First 1
                if (!$row) { throw '找不到這個項目。' }
                $row.name = $name
            } else {
                $id = [Guid]::NewGuid().ToString('N')
                $row = [pscustomobject]@{id=$id; name=$name}
                if ($collection -eq 'Sections') {
                    if ($payload.projectId -notin $data.Projects.id) { throw '找不到所屬專案。' }
                    $row | Add-Member NoteProperty projectId ([string]$payload.projectId)
                }
                $data[$collection] += $row
            }
        }
        'event' {
            $status = 'active'
            if ($id) {
                $existing = $data.Events | Where-Object id -eq $id | Select-Object -First 1
                if ($existing) { $status = $existing.status }
            }
            if ($payload.PSObject.Properties['status']) { $status = [string]$payload.status }
            if ($status -cnotin @('active','completed')) { throw '事件狀態須為進行中或已結束。' }
            if ($payload.sectionId -notin $data.Sections.id) { throw '請先建立分頁。' }
            $title = ([string]$payload.title).Trim()
            if (!$title -or $title.Length -gt 200 -or !(Check-Date $payload.date)) { throw '請輸入 200 字以內的標題與有效日期。' }
            $schedules = @()
            if ($payload.PSObject.Properties['schedules']) {
                if ($null -eq $payload.schedules -or $payload.schedules -isnot [System.Array]) { throw '時程必須是陣列。' }
                foreach ($schedule in $payload.schedules) {
                    $schedules += Convert-Schedule $schedule
                }
                if ($schedules.Count -and @($schedules.id | Select-Object -Unique).Count -ne $schedules.Count) { throw '時程識別碼重複。' }
            } elseif ($payload.start -or $payload.end) {
                if (!(Check-Date $payload.start) -or !(Check-Date $payload.end) -or $payload.end -lt $payload.start) { throw '時程日期無效。' }
                $schedules = @(Convert-Schedule ([pscustomobject]@{name='原有時程'; start=[string]$payload.start; end=[string]$payload.end}))
            }
            $start = ''; $end = ''
            if ($schedules.Count) {
                $start = ($schedules | Sort-Object start | Select-Object -First 1).start
                $end = ($schedules | Sort-Object end | Select-Object -Last 1).end
            }
            $html = [string]$payload.html
            if ($html.Length -gt 500000) { throw '本文超過 50 萬字元。' }
            $files = @($payload.files | Where-Object { $_ })
            Validate-Files $files
            if ($id -and $id -notin $data.Events.id) { throw '找不到這筆事件。' }
            if (!$id) { $id = [Guid]::NewGuid().ToString('N') }
            $data.Events = @($data.Events | Where-Object id -ne $id)
            $data.Events += [pscustomobject]@{
                id=$id; sectionId=[string]$payload.sectionId; title=$title; date=[string]$payload.date
                start=$start; end=$end; schedules=$schedules; html=$html; files=$files; status=$status
            }
        }
        default { throw '未知的資料類型。' }
    }
    Save-Data $data
    Remove-UnusedAttachments $beforeFiles $data
    return $id
}

function Get-AttachmentIds($data) {
    foreach ($event in $data.Events) {
        foreach ($file in $event.files) { if ($file.id) { [string]$file.id } }
        foreach ($schedule in $event.schedules) {
            foreach ($file in $schedule.files) { if ($file.id) { [string]$file.id } }
        }
        foreach ($match in [regex]::Matches([string]$event.html, '/attachment/([a-f0-9]{32}\.(?:png|jpg|gif|webp|bin))')) { $match.Groups[1].Value }
    }
}

function Remove-UnusedAttachments($beforeFiles, $data) {
    $retained = @(Get-AttachmentIds $data)
    foreach ($id in @($beforeFiles | Select-Object -Unique)) {
        if ($id -notmatch '^[a-f0-9]{32}\.(png|jpg|gif|webp|bin)$' -or $id -in $retained) { continue }
        try { [IO.File]::Delete((Join-Path $attachments $id)) }
        catch { $script:cleanupWarnings += "資料已儲存、原檔未刪除：$id" }
    }
}

function Move-Event($payload) {
    $data = Read-Data
    $event = $data.Events | Where-Object id -eq ([string]$payload.id) | Select-Object -First 1
    $target = $data.Sections | Where-Object id -eq ([string]$payload.sectionId) | Select-Object -First 1
    if (!$event -or !$target) { throw '找不到事件或目標分頁。' }
    $source = $data.Sections | Where-Object id -eq $event.sectionId | Select-Object -First 1
    if (!$source -or $source.projectId -ne $target.projectId) { throw '只能搬移至同一專案的分頁。' }
    if ($event.sectionId -ne $target.id) { $event.sectionId = $target.id; Save-Data $data }
    return $event.id
}

function Remove-Container($payload, $kind) {
    $data = Read-Data
    $beforeFiles = @(Get-AttachmentIds $data)
    $id = [string]$payload.id
    $sections = @($id)
    if ($kind -eq 'project') {
        if (!$id -or $id -notin $data.Projects.id) { throw '找不到專案。' }
        $sections = @($data.Sections | Where-Object projectId -eq $id | ForEach-Object id)
        $data.Projects = @($data.Projects | Where-Object id -ne $id)
    } elseif (!$id -or $id -notin $data.Sections.id) { throw '找不到分頁。' }
    $data.Sections = @($data.Sections | Where-Object { $_.id -notin $sections })
    $data.Events = @($data.Events | Where-Object { $_.sectionId -notin $sections })
    Save-Data $data
    Remove-UnusedAttachments $beforeFiles $data
    return $id
}

function Update-EventStatus($payload) {
    if ([string]$payload.status -cnotin @('active','completed')) { throw '事件狀態須為進行中或已結束。' }
    $data = Read-Data
    $event = $data.Events | Where-Object id -eq ([string]$payload.id) | Select-Object -First 1
    if (!$event) { throw '找不到這筆事件。' }
    $event.status = [string]$payload.status
    Save-Data $data
    return $event.id
}

function Remove-Event($payload) {
    $data = Read-Data
    $beforeFiles = @(Get-AttachmentIds $data)
    $id = [string]$payload.id
    if (!$id -or $id -notin $data.Events.id) { throw '找不到這筆事件。' }
    $data.Events = @($data.Events | Where-Object id -ne $id)
    Save-Data $data
    Remove-UnusedAttachments $beforeFiles $data
    return $id
}

function Update-Schedule($payload) {
    $data = Read-Data
    $beforeFiles = @(Get-AttachmentIds $data)
    $event = $data.Events | Where-Object id -eq ([string]$payload.eventId) | Select-Object -First 1
    if (!$event) { throw '找不到這筆事件。' }
    $schedule = Convert-Schedule $payload.schedule
    $index = -1
    for ($s = 0; $s -lt $event.schedules.Count; $s++) {
        if ($event.schedules[$s].id -eq $schedule.id) { $index = $s; break }
    }
    if ($index -lt 0) { throw '找不到這條時程，請重新開啟事件。' }
    $event.schedules[$index] = $schedule
    $event.start = ($event.schedules | Sort-Object start | Select-Object -First 1).start
    $event.end = ($event.schedules | Sort-Object end | Select-Object -Last 1).end
    Save-Data $data
    Remove-UnusedAttachments $beforeFiles $data
    return $schedule.id
}

$web = Join-Path $PSScriptRoot 'web'
$requiredFiles = @(
    'index.html', 'app.css', 'app.js', 'features.js',
    'vendor/pdf-lib.min.js', 'vendor/pdf.mjs', 'vendor/pdf.worker.mjs', 'vendor/xlsx.full.min.js',
    'vendor/tesseract.min.js', 'vendor/worker.min.js',
    'vendor/core/tesseract-core-lstm.wasm.js', 'vendor/core/tesseract-core-simd-lstm.wasm.js',
    'vendor/lang/chi_tra.traineddata.gz', 'vendor/lang/eng.traineddata.gz'
)
foreach ($relativePath in $requiredFiles) {
    if (!(Test-Path -LiteralPath (Join-Path $web $relativePath) -PathType Leaf)) {
        throw "缺少 web/$relativePath。請確認 web 資料夾完整，並放在腳本旁邊。"
    }
}
Import-Module ImportExcel -ErrorAction Stop

[IO.Directory]::CreateDirectory($DataRoot) | Out-Null
$lock = [IO.File]::Open((Join-Path $DataRoot '.session.lock'),'OpenOrCreate','ReadWrite','None')
$listener = $null
$pendingClients = [Collections.Generic.List[object]]::new()
try {
    [IO.Directory]::CreateDirectory($attachments) | Out-Null
    if (Test-Path -LiteralPath $workbook) { $null = Read-Data }
    else { Save-Data (Read-Data) }
    $types = @{
        '.html'='text/html; charset=utf-8'; '.css'='text/css; charset=utf-8'
        '.js'='text/javascript; charset=utf-8'; '.mjs'='text/javascript; charset=utf-8'
        '.wasm'='application/wasm'; '.gz'='application/octet-stream'
        '.png'='image/png'; '.jpg'='image/jpeg'; '.gif'='image/gif'; '.webp'='image/webp'
    }
    $csp = "Content-Security-Policy: default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline'; img-src 'self' blob: data:; worker-src 'self' blob:; connect-src 'self' blob: data:; object-src 'none'; frame-ancestors 'none'; base-uri 'none'`r`n"
    $listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback,0)
    $listener.Start()
    $port = $listener.LocalEndpoint.Port
    $hostHeader = "127.0.0.1:$port"
    $origin = "http://$hostHeader"
    $ip = @()
    try {
        $ip = @([Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | Where-Object OperationalStatus -eq Up | ForEach-Object {
        $_.GetIPProperties().UnicastAddresses
    } | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' -and ![Net.IPAddress]::IsLoopback($_.Address) } | ForEach-Object { $_.Address.ToString() } | Sort-Object -Unique)
    } catch {
        try { $ip = @([Net.Dns]::GetHostAddresses([Net.Dns]::GetHostName()) | Where-Object AddressFamily -eq InterNetwork | ForEach-Object ToString) } catch {}
    }
    Write-Host "Local Data Center: $origin"
    Write-Host "Excel: $workbook"
    if (!$NoBrowser) {
        try { Start-Process $origin } catch { Write-Warning '請手動在瀏覽器開啟上方網址。' }
    }
    while ($true) {
        # Chromium opens idle connections in advance. Do not block the single
        # request loop waiting for an idle socket while worker requests queue up.
        while ($listener.Pending()) {
            $pendingClients.Add(@{client=$listener.AcceptTcpClient(); accepted=[datetime]::UtcNow})
        }
        $ready = $null
        foreach ($pending in @($pendingClients.ToArray())) {
            try {
                if ($pending.client.Available -gt 0 -and !$ready) { $ready = $pending; continue }
                $expired = ([datetime]::UtcNow - $pending.accepted).TotalSeconds -gt 120 -or ($pending.client.Client.Poll(0,[Net.Sockets.SelectMode]::SelectRead) -and $pending.client.Available -eq 0)
            } catch { $expired = $true }
            if ($expired) {
                $pending.client.Dispose()
                $null = $pendingClients.Remove($pending)
            }
        }
        if (!$ready) { Start-Sleep -Milliseconds 10; continue }
        $null = $pendingClients.Remove($ready)
        $client = $ready.client
        $client.NoDelay = $true
        $client.Client.LingerState = [Net.Sockets.LingerOption]::new($true,5)
        $client.ReceiveTimeout = 8000
        $client.SendTimeout = 8000
        $stream = $client.GetStream()
        $script:reuseClient = $false
        try {
            $request = Read-Request $stream
            if (!$request) { continue }
            $path = $request.Path
            $script:cleanupWarnings = @()
            if ($request.Headers.Host -ne $hostHeader) {
                Send-Json $stream @{error='Host 不符。'} '403 Forbidden'
                continue
            }
            if ($request.Method -eq 'POST') {
                if ($request.Headers.Origin -ne $origin -or $request.Headers['X-Local-Token'] -cne $token) {
                    Send-Json $stream @{error='請重新開啟本機網頁。'} '403 Forbidden'
                    continue
                }
                if ($path -eq '/api/save') {
                    $bytes = Read-Body $stream $request 4000000
                    $payload = $utf8.GetString($bytes) | ConvertFrom-Json
                    $savedId = Update-Data $payload
                    Send-Json $stream @{id=$savedId; warnings=@($script:cleanupWarnings)}
                    continue
                }
                if ($path -eq '/api/status') {
                    $bytes = Read-Body $stream $request 4096
                    $payload = $utf8.GetString($bytes) | ConvertFrom-Json
                    Send-Json $stream @{id=(Update-EventStatus $payload)}
                    continue
                }
                if ($path -eq '/api/delete-event') {
                    $bytes = Read-Body $stream $request 4096
                    $payload = $utf8.GetString($bytes) | ConvertFrom-Json
                    $savedId = Remove-Event $payload
                    Send-Json $stream @{id=$savedId; warnings=@($script:cleanupWarnings)}
                    continue
                }
                if ($path -eq '/api/schedule') {
                    $bytes = Read-Body $stream $request 4000000
                    $payload = $utf8.GetString($bytes) | ConvertFrom-Json
                    $savedId = Update-Schedule $payload
                    Send-Json $stream @{id=$savedId; warnings=@($script:cleanupWarnings)}
                    continue
                }
                if ($path -in @('/api/move-event','/api/delete-project','/api/delete-section')) {
                    $payload = $utf8.GetString((Read-Body $stream $request 4096)) | ConvertFrom-Json
                    if ($path -eq '/api/move-event') { $savedId = Move-Event $payload }
                    elseif ($path -eq '/api/delete-project') { $savedId = Remove-Container $payload 'project' }
                    else { $savedId = Remove-Container $payload 'section' }
                    Send-Json $stream @{id=$savedId; warnings=@($script:cleanupWarnings)}
                    continue
                }
                if ($path -eq '/api/upload') {
                    $bytes = Read-Body $stream $request 25000000
                    $ext = 'bin'
                    switch ($request.Headers['X-Image-Type']) {
                        'image/png' { $ext = 'png' }
                        'image/jpeg' { $ext = 'jpg' }
                        'image/gif' { $ext = 'gif' }
                        'image/webp' { $ext = 'webp' }
                    }
                    $id = [Guid]::NewGuid().ToString('N') + '.' + $ext
                    [IO.File]::WriteAllBytes((Join-Path $attachments $id),$bytes)
                    Send-Json $stream @{id=$id}
                    continue
                }
            }
            if ($request.Method -eq 'GET') {
                if ($path -eq '/api/data') {
                    $data = Read-Data
                    $data.apiVersion = 3
                    $data.capabilities = @('event-delete','attachment-cleanup','event-move','container-delete','rich-tables')
                    $data.ip = $ip
                    Send-Json $stream $data
                    continue
                }
                if ($path -eq '/') {
                    $html = [IO.File]::ReadAllText((Join-Path $web 'index.html'),$utf8).Replace('__LOCAL_TOKEN__',$token)
                    Send-Response $stream '200 OK' $types['.html'] ($utf8.GetBytes($html)) $csp
                    continue
                }
                if ($path -match '^/attachment/([a-f0-9]{32}\.(png|jpg|gif|webp|bin))$') {
                    $name = $Matches[1]
                    $file = Join-Path $attachments $name
                    if ([IO.File]::Exists($file)) {
                        $type = $types[[IO.Path]::GetExtension($name)]
                        $extra = ''
                        if (!$type) { $type = 'application/octet-stream'; $extra = "Content-Disposition: attachment`r`n" }
                        Send-Response $stream '200 OK' $type ([IO.File]::ReadAllBytes($file)) $extra
                        continue
                    }
                    Send-Json $stream @{error="找不到附件原檔：$file。請確認備份的 Attachments 已放回這個資料位置，並保留原始檔名。"} '404 Not Found'
                    continue
                }
                if ($path -match '^/(app\.(js|css)|features\.js|vendor/[A-Za-z0-9_.-]+(?:/[A-Za-z0-9_.-]+)*)$' -and $path -notmatch '(^|/)\.{1,2}(/|$)') {
                    $file = Join-Path $web $path.Substring(1)
                    if ([IO.File]::Exists($file)) {
                        $type = $types[[IO.Path]::GetExtension($file)]
                        if (!$type) { $type = 'application/octet-stream' }
                        $cacheControl = 'no-store'
                        if ($path.StartsWith('/vendor/')) { $cacheControl = 'private, max-age=86400' }
                        Send-Response $stream '200 OK' $type ([IO.File]::ReadAllBytes($file)) $csp $cacheControl
                        continue
                    }
                }
            }
            Send-Json $stream @{error='找不到此功能。'} '404 Not Found'
        } catch {
            try { Send-Json $stream @{error=$_.Exception.Message} '400 Bad Request' } catch {}
        } finally {
            if ($script:reuseClient) { $pendingClients.Add(@{client=$client; accepted=[datetime]::UtcNow}) }
            else {
                try { $client.Client.Shutdown([Net.Sockets.SocketShutdown]::Send) } catch {}
                $client.Dispose()
            }
        }
    }
} finally {
    foreach ($pending in $pendingClients) { $pending.client.Dispose() }
    if ($listener) { $listener.Stop() }
    $lock.Dispose()
}
