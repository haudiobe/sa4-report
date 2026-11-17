<#
SA4 local listing agent (after T-2026.10.27; docs/LOCAL_LISTING_AGENT.md).

Runs on ONE Windows laptop on the meeting-room network, started every two minutes by Windows
Task Scheduler (install-scheduled-task.ps1). One run scans every working group once:

  GET <baseUrl><group>/          plain HTTP from this laptop (no browser: no CORS, no mixed content)
  checks                         HTTP 200, HTML, the response came from the configured host,
                                 the page is the listing of THAT group's folder, it has links
  <outDir>\<group>.listing.json  only the listing's links (and their names), with its checksum --
                                 never a TDoc file; written atomically (temp file + rename)
  <outDir>\<group>.status.json   last attempt, last success, error (informative)

<outDir> is a folder that Google Drive for desktop synchronizes; the reports read it. A failed
scan never touches the last good listing. Nothing is sent anywhere else; the meeting-room server
is only read.

  .\sa4-local-listing-agent.ps1 -ConfigPath .\agent.config.json         (one scan of all groups)
  .\sa4-local-listing-agent.ps1 -ConfigPath .\agent.config.json -Group MBS
  .\sa4-local-listing-agent.ps1 -ConfigPath .\agent.config.json -IgnoreWindow     (scan now, even overnight)
  .\sa4-local-listing-agent.ps1 -ConfigPath .\agent.config.json -CheckWindow [-AtUtc <time>]
                                 (only say whether that moment is inside the window; no scan, no file written)
Windows PowerShell 5.1 or PowerShell 7.

Operating window: scans only from 07:30 to 19:30 in the meeting's local time (Calgary,
'Mountain Standard Time', daylight saving included), whatever time zone the laptop is set to.
Configurable with "operatingWindow": { "timeZone", "start", "end" } in agent.config.json. A run
outside the window is a normal pause, not an error: one log line, exit code 0, and neither the
listing nor the status files are touched.
#>
param(
  [Parameter(Mandatory = $true)][string]$ConfigPath,
  [string]$Group,
  [switch]$IgnoreWindow,
  [switch]$CheckWindow,
  [datetime]$AtUtc
)
$ErrorActionPreference = 'Stop'
$cfg = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$baseUrl = ([string]$cfg.baseUrl).TrimEnd('/') + '/'
$outDir = [string]$cfg.outDir
$groups = @($cfg.groups)
if ($Group) { $groups = @($Group) }
$timeout = 15
if ($cfg.timeoutSeconds) { $timeout = [int]$cfg.timeoutSeconds }
$logDir = Join-Path $env:LOCALAPPDATA 'SA4LocalListingAgent'
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logPath = Join-Path $logDir 'agent.log'

function Write-Log([string]$Text) {
  if ((Test-Path -LiteralPath $logPath) -and ((Get-Item -LiteralPath $logPath).Length -gt 1MB)) { Move-Item -LiteralPath $logPath -Destination "$logPath.1" -Force }
  $line = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') + '  ' + $Text
  Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
  Write-Host $line
}

function Get-Sha256Hex([string]$Text) {
  $sha = [System.Security.Cryptography.SHA256]::Create()
  try { return (($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text)) | ForEach-Object { $_.ToString('x2') }) -join '') } finally { $sha.Dispose() }
}

function Write-Atomic([string]$Path, [string]$Text) {
  $dir = Split-Path -Parent $Path
  $tmp = Join-Path $dir ('.' + (Split-Path -Leaf $Path) + '.tmp')
  [System.IO.File]::WriteAllText($tmp, $Text, (New-Object System.Text.UTF8Encoding($false)))
  Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function ConvertTo-HtmlText([string]$Text) { return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;') }

# Only the links of a listing: '<a href="...">name</a>' per line, in the order of the page.
function Get-ListingLinks([string]$Html) {
  $links = New-Object System.Collections.Generic.List[string]
  $re = [regex]'(?is)<a\s[^>]*?href\s*=\s*["'']([^"''<>]+)["''][^>]*>(.*?)</a>'
  foreach ($m in $re.Matches($Html)) {
    $href = $m.Groups[1].Value.Trim()
    $name = [System.Net.WebUtility]::HtmlDecode(([regex]::Replace($m.Groups[2].Value, '<[^>]*>', ' ') -replace '\s+', ' ').Trim())
    if (-not $href -or -not $name) { continue }
    $links.Add('<a href="' + $href.Replace('"', '%22') + '">' + (ConvertTo-HtmlText $name) + '</a>')
  }
  return , $links
}

function Invoke-GroupScan([string]$G) {
  $folder = $baseUrl + $G + '/'
  $now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
  $statusPath = Join-Path $outDir "$G.status.json"
  $previous = $null
  if (Test-Path -LiteralPath $statusPath) { try { $previous = Get-Content -LiteralPath $statusPath -Raw | ConvertFrom-Json } catch { $previous = $null } }
  $lastSuccess = $null; $lastWritten = $null
  if ($previous) { $lastSuccess = $previous.lastSuccessAt; $lastWritten = $previous.writtenAt }
  # Stages kept apart: scanOk (the LAN read worked), writtenAt (the listing file was written locally). Whether Google Drive
  # has uploaded it is not known here; the report judges that by what it finds in Drive.
  $status = [ordered]@{ format = 'sa4-local-listing-status'; group = $G; folder = $folder; fetchedAt = $now; ok = $false; scanOk = $false; lastSuccessAt = $lastSuccess; writtenAt = $lastWritten; httpStatus = $null; links = 0; error = $null; agent = $env:COMPUTERNAME }
  try {
    $r = Invoke-WebRequest -Uri $folder -UseBasicParsing -TimeoutSec $timeout -MaximumRedirection 0 -Headers @{ 'Cache-Control' = 'no-cache' }
    $status.httpStatus = [int]$r.StatusCode
    if ($status.httpStatus -ne 200) { throw "HTTP $($status.httpStatus)" }
    $type = [string]$r.Headers['Content-Type']
    if ($type -notmatch '(?i)text/html') { throw "not an HTML page ($type)" }
    $body = [string]$r.Content
    $path = ([uri]$folder).AbsolutePath
    if ($body.IndexOf($path, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { throw "the page is not the listing of $path" }
    $links = Get-ListingLinks $body
    if ($links.Count -eq 0) { throw 'no links in the listing' }
    $payload = ($links -join "`n")
    $sha = Get-Sha256Hex $payload
    $status.scanOk = $true; $status.lastSuccessAt = $now; $status.links = $links.Count
    $listing = [ordered]@{ format = 'sa4-local-listing'; version = 1; group = $G; folder = $folder; fetchedAt = $now; scanId = $sha; agent = $env:COMPUTERNAME; linkCount = $links.Count; sha256 = $sha; payload = $payload }
    try { Write-Atomic (Join-Path $outDir "$G.listing.json") ((New-Object PSObject -Property $listing) | ConvertTo-Json -Depth 3) } catch { throw ('write failed: ' + $_.Exception.Message) }
    $status.ok = $true; $status.writtenAt = $now
  } catch {
    $status.error = $_.Exception.Message
    try { if ($_.Exception.Response) { $status.httpStatus = [int]$_.Exception.Response.StatusCode } } catch { }
  }
  Write-Atomic $statusPath ((New-Object PSObject -Property $status) | ConvertTo-Json)
  if ($status.ok) { Write-Log ("{0,-10} ok, {1} links" -f $G, $status.links) } else { Write-Log ("{0,-10} FAILED: {1} (last good listing kept)" -f $G, $status.error) }
  return $status.ok
}

# Operating window, in the meeting's local time. Outside it: a normal pause (exit 0), no file touched.
$window = [ordered]@{ timeZone = 'Mountain Standard Time'; start = '07:30'; end = '19:30' }
if ($cfg.operatingWindow) {
  foreach ($k in @('timeZone', 'start', 'end')) { if ($cfg.operatingWindow.$k) { $window[$k] = [string]$cfg.operatingWindow.$k } }
}
if (-not $IgnoreWindow -or $CheckWindow) {
  try {
    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($window.timeZone)
    $start = [timespan]::Parse($window.start); $end = [timespan]::Parse($window.end)
    if ($start -ge $end) { throw 'start must be before end' }
  } catch {
    Write-Log ("scan not run: operatingWindow is invalid ({0} {1}-{2}): {3}" -f $window.timeZone, $window.start, $window.end, $_.Exception.Message)
    exit 2
  }
  $utcNow = [DateTime]::UtcNow
  if ($PSBoundParameters.ContainsKey('AtUtc')) { $utcNow = [DateTime]::SpecifyKind($AtUtc, [DateTimeKind]::Utc) }
  $meetingNow = [System.TimeZoneInfo]::ConvertTimeFromUtc($utcNow, $tz)
  $inside = ($meetingNow.TimeOfDay -ge $start -and $meetingNow.TimeOfDay -lt $end)
  $when = '{0} {1}' -f $meetingNow.ToString('yyyy-MM-dd HH:mm'), $window.timeZone
  if ($CheckWindow) {
    if ($inside) { $verdict = 'inside' } else { $verdict = 'outside' }
    Write-Host ("{0}Z = {1}: {2} the operating window {3}-{4}" -f $utcNow.ToString('yyyy-MM-ddTHH:mm'), $when, $verdict, $window.start, $window.end)
    exit 0
  }
  if (-not $inside) {
    Write-Log ("paused: {0} is outside the operating window {1}-{2}; scanning resumes at {1} (last listings kept)" -f $when, $window.start, $window.end)
    exit 0
  }
}

# One run at a time (a scheduled run that overlaps a slow one is skipped).
$mutex = New-Object System.Threading.Mutex($false, 'Global\SA4LocalListingAgent')
if (-not $mutex.WaitOne(0)) { Write-Log 'another scan is still running; skipped'; exit 0 }
try {
  if (-not (Test-Path -LiteralPath $outDir)) { throw "output folder $outDir does not exist (is Google Drive for desktop running?)" }
  $failed = 0
  foreach ($g in $groups) { if (-not (Invoke-GroupScan $g)) { $failed++ } }
  if ($failed -gt 0) { exit 1 } else { exit 0 }
} catch {
  Write-Log ('scan not run: ' + $_.Exception.Message)
  exit 2
} finally {
  $mutex.ReleaseMutex() | Out-Null
  $mutex.Dispose()
}
