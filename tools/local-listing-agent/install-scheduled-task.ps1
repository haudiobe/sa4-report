<#
Registers (or removes) the Windows Task Scheduler task that runs the SA4 local listing agent
every two minutes during the operating window (default 07:30-19:30 Calgary time, daily), as the
signed-in user (Google Drive for desktop must be running in that session), only while that user
is logged on. No administrator rights, no stored password.

  .\install-scheduled-task.ps1 -ConfigPath C:\SA4Agent\agent.config.json
  .\install-scheduled-task.ps1 -Uninstall

The window comes from "operatingWindow" in the config (timeZone, start, end; defaults
'Mountain Standard Time', 07:30, 19:30), the same values the agent itself enforces.
  - Laptop clock set to that time zone: a daily trigger at <start> local time, repeating every
    <IntervalMinutes> for (end - start). Daylight saving is followed by Windows.
  - Laptop clock in another time zone: a daily trigger fixed in UTC ("synchronize across time
    zones"), starting one hour before <start> and lasting two hours longer, so that neither a
    later change of the laptop's time zone nor a daylight-saving change in the meeting's zone
    can cut the window short. The agent pauses the extra runs (exit 0, no file touched).
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

# Operating window (same defaults as the agent).
$cfg = Get-Content -LiteralPath $config -Raw | ConvertFrom-Json
$window = [ordered]@{ timeZone = 'Mountain Standard Time'; start = '07:30'; end = '19:30' }
if ($cfg.operatingWindow) {
  foreach ($k in @('timeZone', 'start', 'end')) { if ($cfg.operatingWindow.$k) { $window[$k] = [string]$cfg.operatingWindow.$k } }
}
$tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($window.timeZone)
$start = [timespan]::Parse($window.start); $end = [timespan]::Parse($window.end)
if ($start -ge $end) { throw "operatingWindow: start ($($window.start)) must be before end ($($window.end))." }

$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $script + '" -ConfigPath "' + $config + '"')
$sameZone = ((Get-TimeZone).Id -eq $tz.Id)
if ($sameZone) {
  $first = (Get-Date).Date.Add($start)
  $duration = $end - $start
  $trigger = New-ScheduledTaskTrigger -Daily -At $first
  $trigger.StartBoundary = $first.ToString("yyyy-MM-dd'T'HH:mm:ss")   # no UTC offset: local time, follows daylight saving
  $mode = "daily at $($window.start) local time ($($tz.Id)) for $([math]::Floor($duration.TotalHours))h $($duration.Minutes)min"
} else {
  $meetingToday = [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz).Date
  $firstLocal = $meetingToday.Add($start).AddHours(-1)
  $offset = $tz.GetUtcOffset($firstLocal)
  $firstFixed = New-Object System.DateTimeOffset($firstLocal, $offset)
  $duration = ($end - $start) + (New-TimeSpan -Hours 2)
  $trigger = New-ScheduledTaskTrigger -Daily -At $firstFixed.LocalDateTime
  $trigger.StartBoundary = $firstFixed.ToString("yyyy-MM-dd'T'HH:mm:sszzz")
  $mode = "daily from $($firstFixed.ToString('HH:mm zzz')) (fixed in UTC) for $([math]::Floor($duration.TotalHours))h $($duration.Minutes)min; the laptop is on '$((Get-TimeZone).Id)', not '$($tz.Id)', so the trigger is one hour wider on each side and the agent pauses runs outside $($window.start)-$($window.end) $($tz.Id)"
}
$trigger.Repetition = (New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) -RepetitionDuration $duration).Repetition
$trigger.Repetition.StopAtDurationEnd = $false   # a scan already running at the end of the window finishes (5-minute limit)
$principal = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Write-Host "Registered '$TaskName': every $IntervalMinutes minutes, $mode, while you are logged on."
Write-Host "Scanning window: $($window.start)-$($window.end) $($tz.Id). Log: $env:LOCALAPPDATA\SA4LocalListingAgent\agent.log"
