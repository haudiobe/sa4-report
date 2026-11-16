<#
Registers (or removes) the Windows Task Scheduler task that runs the SA4 local listing agent
every two minutes, as the signed-in user (Google Drive for desktop must be running in that
session), only while that user is logged on. No administrator rights, no stored password.

  .\install-scheduled-task.ps1 -ConfigPath C:\SA4Agent\agent.config.json
  .\install-scheduled-task.ps1 -Uninstall
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
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $script + '" -ConfigPath "' + $config + '"')
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) -RepetitionDuration (New-TimeSpan -Days 30)
$principal = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Write-Host "Registered '$TaskName': every $IntervalMinutes minutes while you are logged on. Log: $env:LOCALAPPDATA\SA4LocalListingAgent\agent.log"
