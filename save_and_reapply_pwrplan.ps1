# Save the current power plan as a new power scheme
$PowerSchemeName = "MyCustomPowerPlan"
Powercfg -EXPORT -DESTINATION "$env:USERPROFILE\Desktop" -SCHEME $PowerSchemeName

# Get the GUID of the new power scheme
$PowerSchemeGUID = (Powercfg -LIST | Select-String -Pattern "$PowerSchemeName").ToString().Split(" ")[-1]

# Set the new power scheme as the active one
Powercfg -SETACTIVE $PowerSchemeGUID

# Create a scheduled task to apply the power plan after reboot
$TaskName = "ApplyPowerPlan"
$TaskAction = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument "-Command &{Powercfg -SETACTIVE $PowerSchemeGUID}"
$TaskTrigger1 = New-ScheduledTaskTrigger -AtStartup
$TaskTrigger2 = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Hours 1)
Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger1,$TaskTrigger2 -Description "Applies the custom power plan" -User $env:USERNAME -Force

# Optional: Add a task to re-apply the power plan every hour
# $TaskNameHourly = "ApplyPowerPlanHourly"
# $TaskActionHourly = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument "-Command &{Powercfg -SETACTIVE $PowerSchemeGUID}"
# $TaskTriggerHourly = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Hours 1)
# Register-ScheduledTask -TaskName $TaskNameHourly -Action $TaskActionHourly -Trigger $TaskTriggerHourly -Description "Applies the custom power plan every hour" -User $env:USERNAME -Force
