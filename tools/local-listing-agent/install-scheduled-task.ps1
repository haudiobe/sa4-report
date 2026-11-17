<#
Registers (or removes) the Windows Task Scheduler task that runs the SA4 local listing agent
every two minutes during the operating window (default 07:30-19:30 Calgary time, ABT, daily), as the
signed-in user (Google Drive for desktop must be running in that session), only while that user
is logged on. No administrator rights, no stored password.

  .\install-scheduled-task.ps1 -ConfigPath C:\SA4Agent\agent.config.json
  .\install-scheduled-task.ps1 -Uninstall

The window comes from "operatingWindow" in the config, resolved exactly as the agent does
(Resolve-OperatingWindow; default 07:30-19:30 at UTC-06:00, ABT).
  - A fixed utcOffset (the default): a daily trigger fixed in UTC ("synchronize across time
    zones") at <start>, repeating every <IntervalMinutes> for (end - start). Neither the laptop's
    time zone nor its Windows time-zone rules matter.
  - A Windows timeZone, laptop clock set to that zone: a daily trigger at <start> local time for
    (end - start); daylight saving is followed by Windows.
  - A Windows timeZone, laptop clock in another zone: fixed in UTC, one hour before <start> and
    two hours longer, so that neither a later change of the laptop's zone nor a daylight-saving
    change can cut the window short. The agent pauses the extra runs (exit 0, no file touched).
#>
param(
  [string]$ConfigPath,
  [string]$TaskName = 'SA4 local listing agent',
  [int]$IntervalMinutes = 2,
  [switch]$Uninstall
)
$ErrorActionPreference = 'Stop'
if ($Uninstall) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
  Write-Host "Removed the task '$TaskName'."
  return
}
if (-not $ConfigPath) { throw '-ConfigPath is required.' }
$config = (Resolve-Path -LiteralPath $ConfigPath).Path
$script = Join-Path $PSScriptRoot 'sa4-local-listing-agent.ps1'
if (-not (Test-Path -LiteralPath $script)) { throw "Not found: $script" }

# Operating window, from "operatingWindow" in the config. Default: 07:30-19:30 at UTC-06:00 (ABT), Calgary's time
# all year from 1 November 2026 (docs/MEETING_TIME_ZONES.md). A fixed offset does not depend on the laptop's Windows
# time-zone rules, which may predate Alberta's change ('Mountain Standard Time' there still goes to UTC-7 in winter).
# "timeZone" (a Windows zone id) without "utcOffset" uses that zone's Windows rules instead, e.g. for a meeting elsewhere.
function Resolve-OperatingWindow($Config) {
  $w = [ordered]@{ utcOffset = '-06:00'; label = 'ABT'; timeZone = ''; start = '07:30'; end = '19:30' }
  $o = $Config.operatingWindow
  if ($o) {
    if ($o.timeZone -and -not $o.utcOffset) { $w.utcOffset = ''; $w.label = '' }
    foreach ($k in @('utcOffset', 'label', 'timeZone', 'start', 'end')) { if ($o.$k) { $w[$k] = [string]$o.$k } }
  }
  if ($w.utcOffset) {
    if ($w.utcOffset -notmatch '^[+-]\d{2}:\d{2}$') { throw "utcOffset must look like -06:00 (it is '$($w.utcOffset)')" }
    $offset = [timespan]::Parse($w.utcOffset.TrimStart('+'))
    $name = 'UTC' + $w.utcOffset
    if ($w.label) { $name = $w.label + ' (' + $name + ')' }
    $tz = [System.TimeZoneInfo]::CreateCustomTimeZone('SA4 meeting time', $offset, $name, $name)
    $fixed = $true
  } else {
    $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($w.timeZone)
    $name = $w.timeZone
    $fixed = $false
  }
  $start = [timespan]::Parse($w.start); $end = [timespan]::Parse($w.end)
  if ($start -ge $end) { throw "start ($($w.start)) must be before end ($($w.end))" }
  return [pscustomobject]@{ tz = $tz; name = $name; fixed = $fixed; start = $start; end = $end; startText = $w.start; endText = $w.end; text = "$($w.start)-$($w.end) $name" }
}

$cfg = Get-Content -LiteralPath $config -Raw | ConvertFrom-Json
$window = Resolve-OperatingWindow $cfg
$tz = $window.tz; $start = $window.start; $end = $window.end

$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $script + '" -ConfigPath "' + $config + '"')
$sameZone = (-not $window.fixed) -and ((Get-TimeZone).Id -eq $tz.Id)
if ($window.fixed) {
  $meetingToday = [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz).Date
  $firstFixed = New-Object System.DateTimeOffset($meetingToday.Add($start), $tz.BaseUtcOffset)
  $duration = $end - $start
  $trigger = New-ScheduledTaskTrigger -Daily -At $firstFixed.LocalDateTime
  $trigger.StartBoundary = $firstFixed.ToString("yyyy-MM-dd'T'HH:mm:sszzz")
  $mode = "daily at $($window.startText) $($window.name) (fixed in UTC) for $([math]::Floor($duration.TotalHours))h $($duration.Minutes)min"
} elseif ($sameZone) {
  $first = (Get-Date).Date.Add($start)
  $duration = $end - $start
  $trigger = New-ScheduledTaskTrigger -Daily -At $first
  $trigger.StartBoundary = $first.ToString("yyyy-MM-dd'T'HH:mm:ss")   # no UTC offset: local time, follows daylight saving
  $mode = "daily at $($window.startText) local time ($($tz.Id)) for $([math]::Floor($duration.TotalHours))h $($duration.Minutes)min"
} else {
  $meetingToday = [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz).Date
  $firstLocal = $meetingToday.Add($start).AddHours(-1)
  $offset = $tz.GetUtcOffset($firstLocal)
  $firstFixed = New-Object System.DateTimeOffset($firstLocal, $offset)
  $duration = ($end - $start) + (New-TimeSpan -Hours 2)
  $trigger = New-ScheduledTaskTrigger -Daily -At $firstFixed.LocalDateTime
  $trigger.StartBoundary = $firstFixed.ToString("yyyy-MM-dd'T'HH:mm:sszzz")
  $mode = "daily from $($firstFixed.ToString('HH:mm zzz')) (fixed in UTC) for $([math]::Floor($duration.TotalHours))h $($duration.Minutes)min; the laptop is on '$((Get-TimeZone).Id)', not '$($tz.Id)', so the trigger is one hour wider on each side and the agent pauses runs outside $($window.text)"
}
$trigger.Repetition = (New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) -RepetitionDuration $duration).Repetition
$trigger.Repetition.StopAtDurationEnd = $false   # a scan already running at the end of the window finishes (5-minute limit)
$principal = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Write-Host "Registered '$TaskName': every $IntervalMinutes minutes, $mode, while you are logged on."
Write-Host "Scanning window: $($window.text). Log: $env:LOCALAPPDATA\SA4LocalListingAgent\agent.log"
