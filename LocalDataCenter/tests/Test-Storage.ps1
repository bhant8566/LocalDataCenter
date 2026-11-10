$ErrorActionPreference = 'Stop'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('ldc-storage-' + [guid]::NewGuid().ToString('N'))
$DataRoot = $testRoot
$attachments = Join-Path $testRoot 'Attachments'
$workbook = Join-Path $testRoot 'LocalDataCenter.xlsx'
$utf8 = [Text.UTF8Encoding]::new($false)
[IO.Directory]::CreateDirectory($attachments) | Out-Null
Import-Module ImportExcel -ErrorAction Stop
$source = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '../LocalDataCenter.ps1'))
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseInput($source,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count) { throw 'PowerShell parse failed' }
foreach ($function in $ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$false)) {
    . ([scriptblock]::Create($function.Extent.Text))
}
$schema = [ordered]@{ Projects=@('id','name'); Sections=@('id','projectId','name'); Events=@('id','sectionId','title','date','start','end'); Content=@('eventId','part','data') }
$project = Update-Data ([pscustomobject]@{kind='project';name='Storage'})
$section = Update-Data ([pscustomobject]@{kind='section';projectId=$project;name='Tests'})
$fileId = [guid]::NewGuid().ToString('N') + '.bin'
$filePath = Join-Path $attachments $fileId
[IO.File]::WriteAllText($filePath,'locked attachment')
$payload = [pscustomobject]@{kind='event';sectionId=$section;title='Locked file';date='2026-11-10';schedules=@();files=@([pscustomobject]@{id=$fileId;name='locked.txt'});html=''}
$eventId = Update-Data $payload
$payload | Add-Member NoteProperty id $eventId
$payload.files = @()
$script:cleanupWarnings = @()
$handle = [IO.File]::Open($filePath,'Open','ReadWrite','None')
try { $null = Update-Data $payload } finally { $handle.Dispose() }
if (!$script:cleanupWarnings.Count -or $script:cleanupWarnings[0] -notlike '*資料已儲存、原檔未刪除*') { throw 'Expected cleanup warning for locked file' }
$stored = Read-Data
if (@($stored.Events[0].files).Count) { throw 'Cleanup failure rolled back saved data' }
if (!(Test-Path -LiteralPath $filePath)) { throw 'Locked attachment disappeared' }
$payload.html = '<img src="/attachment/' + $fileId + '">'
$null = Update-Data $payload
$payload.files = @([pscustomobject]@{id=$fileId;name='locked.txt'})
$null = Update-Data $payload
$payload.files = @()
$null = Update-Data $payload
if (!(Test-Path -LiteralPath $filePath)) { throw 'Inline image reference was not retained' }
$payload.html = ''
$null = Update-Data $payload
if (Test-Path -LiteralPath $filePath) { throw 'Removed inline attachment was not deleted' }
if (!(Test-Path (Join-Path $DataRoot 'LocalDataCenter.previous.xlsx'))) { throw 'Previous workbook backup missing' }
Write-Output "PASS: locked-file warning, committed data, inline references, physical deletion and previous workbook backup. Temporary data: $testRoot"
