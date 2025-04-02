# PowerShell script to redirect Windows Event Logs (Security, System, Application) to D:\Logs\
# Run this script with Administrator privileges
 
# Define the logs to redirect
$logsToRedirect = @("Security", "System", "Application")
 
# Base directory for logs
$baseLogPath = "D:\Logs"
 
# Process each log type
foreach ($logName in $logsToRedirect) {
    # Create the target directory if it doesn't exist
    $logPath = "$baseLogPath\$logName"
    if (-not (Test-Path -Path $logPath)) {
        Write-Host "Creating directory: $logPath"
        New-Item -Path $logPath -ItemType Directory -Force
    }
 
    # Create a backup of the current registry settings
    $backupPath = "$baseLogPath\${logName}_EventLogRegistry_Backup.reg"
    Write-Host "Creating registry backup at: $backupPath"
    reg export "HKLM\SYSTEM\CurrentControlSet\Services\EventLog\$logName" $backupPath /y
 
    # Set registry path for this log
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\$logName"
    # Get current log path
    $currentLogPath = (Get-ItemProperty -Path $regPath -Name "File" -ErrorAction SilentlyContinue).File
    Write-Host "Current $logName log path is: $currentLogPath"
 
    # Update the registry to point to new location
    try {
        Write-Host "Updating registry to redirect $logName log to: $logPath\$logName.evtx"
        Set-ItemProperty -Path $regPath -Name "File" -Value "$logPath\$logName.evtx"
        # Set maximum log size (4GB)
        Set-ItemProperty -Path $regPath -Name "MaxSize" -Value 4294967296
        Write-Host "Registry updated successfully for $logName log."
        # Create a scheduled task to clear this event log after next reboot
        $taskName = "Clear${logName}LogOnce"
        $action = New-ScheduledTaskAction -Execute "wevtutil" -Argument "cl $logName"
        $trigger = New-ScheduledTaskTrigger -AtStartup
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        # Register the task to run once at next startup
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
        Write-Host "Scheduled task '$taskName' created to clear $logName log at next startup."
    } 
    catch {
        Write-Error "Failed to redirect $logName event log: $_"
    }
 
    # Verify the registry change
    $updatedLogPath = (Get-ItemProperty -Path $regPath -Name "File").File
    Write-Host "Updated $logName log path in registry is now: $updatedLogPath"
    Write-Host "----------------------------------------"
}
 
# Create a cleanup scheduled task that removes itself after execution
$cleanupScript = @"
# Remove all the clear event log tasks
Get-ScheduledTask -TaskName "Clear*LogOnce" | Unregister-ScheduledTask -Confirm:`$false
# Remove this cleanup task itself
Unregister-ScheduledTask -TaskName "RemoveEventLogTasks" -Confirm:`$false
"@
 
$cleanupScriptPath = "$baseLogPath\CleanupTasks.ps1"
$cleanupScript | Out-File -FilePath $cleanupScriptPath -Force
 
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$cleanupScriptPath`""
$trigger = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
 
Register-ScheduledTask -TaskName "RemoveEventLogTasks" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
 
Write-Host "IMPORTANT: Changes will take effect after system reboot."
Write-Host "Please reboot the system for all event log redirections to take effect."