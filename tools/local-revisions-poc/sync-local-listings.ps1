<#
DISPOSABLE PROOF OF CONCEPT -- the local listing agent (issue #16 follow-up: automatic local revisions).

Runs on ONE Windows laptop on the meeting-room network. It reads the folder listings of the
meeting-room server (plain HTTP GET, no browser, so no CORS or mixed-content rule applies) and
writes each one into a folder that Google Drive for desktop synchronizes. The reports read
these files with the Drive access they already have; the agent needs no Google credentials of
its own, and the meeting-room server is not changed or exposed.

Per working group it writes, atomically (temp file + rename):
  <OutDir>\<Group>.html          the listing as served, with a first line
                                 <!-- sa4-local-listing folder="<url>" fetched="<ISO UTC>" -->
  <OutDir>\<Group>.status.json   { folder, ok, fetchedAt, lastSuccessAt, httpStatus, bytes, error }
A failed fetch leaves the last good .html as it is and records the error in .status.json.

  .\sync-local-listings.ps1 -OutDir "G:\My Drive\SA4-138 local listings" -Once
  .\sync-local-listings.ps1 -OutDir "G:\My Drive\SA4-138 local listings" -IntervalSeconds 120
Windows PowerShell 5.1 or PowerShell 7. Stop with Ctrl+C (or end the scheduled task).
#>
param(
  [Parameter(Mandatory = $true)][string]$OutDir,
  [string]$BaseUrl = 'http://10.10.10.10/ftp/SA/SA4/Inbox/Drafts/',
  [string[]]$Groups = @('MBS', 'Video', 'FS_6G_MED', 'RTC', 'Audio'),
  [int]$IntervalSeconds = 120,
  [int]$TimeoutSeconds = 10,
  [switch]$Once
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

function Write-Atomic([string]$Path, [string]$Text) {
  $tmp = "$Path.tmp"
  [System.IO.File]::WriteAllText($tmp, $Text, (New-Object System.Text.UTF8Encoding($false)))
  Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function Sync-Group([string]$Group) {
  $folder = $BaseUrl.TrimEnd('/') + '/' + $Group + '/'
  $statusPath = Join-Path $OutDir "$Group.status.json"
  $previous = $null
  if (Test-Path -LiteralPath $statusPath) { try { $previous = Get-Content -LiteralPath $statusPath -Raw | ConvertFrom-Json } catch { $previous = $null } }
  $now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
  $status = [ordered]@{ folder = $folder; ok = $false; fetchedAt = $now; lastSuccessAt = $(if ($previous) { $previous.lastSuccessAt } else { $null }); httpStatus = $null; bytes = 0; error = $null }
  try {
    $r = Invoke-WebRequest -Uri $folder -UseBasicParsing -TimeoutSec $TimeoutSeconds -Headers @{ 'Cache-Control' = 'no-cache' }
    $status.httpStatus = [int]$r.StatusCode
    $body = [string]$r.Content
    if ($body -notmatch '(?i)<a\s') { throw "no links in the response (not a folder listing?)" }
    $header = '<!-- sa4-local-listing folder="' + $folder + '" fetched="' + $now + '" -->'
    Write-Atomic (Join-Path $OutDir "$Group.html") ($header + "`n" + $body)
    $status.ok = $true; $status.lastSuccessAt = $now; $status.bytes = $body.Length
  } catch {
    $status.error = $_.Exception.Message
    if ($_.Exception.Response) { try { $status.httpStatus = [int]$_.Exception.Response.StatusCode } catch { } }
  }
  Write-Atomic $statusPath ((New-Object PSObject -Property $status) | ConvertTo-Json)
  $msg = if ($status.ok) { "ok, $($status.bytes) bytes" } else { "FAILED: $($status.error)" }
  Write-Host ("{0}  {1,-10} {2}" -f $now, $Group, $msg)
}

do {
  foreach ($g in $Groups) { Sync-Group $g }
  if (-not $Once) { Start-Sleep -Seconds $IntervalSeconds }
} while (-not $Once)
