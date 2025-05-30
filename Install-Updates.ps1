# --- BEGIN SCRIPT --- 
# Load required .NET assembly for the File Open dialog 
Add-Type -AssemblyName System.Windows.Forms 

# Open a File Explorer window to select the CSV file containing the server list. 
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
$openFileDialog.Filter = "CSV Files (*.csv)|*.csv" 
if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { 
    Write-Host "No CSV file selected. Exiting..." 
    exit 
} 
$csvPath = $openFileDialog.FileName 
$servers = Import-Csv -Path $csvPath 

# Function: Copy update files to each server (using the administrative share on drive D) 
function Copy-UpdatesFiles { 
    param ( 
        [string]$sourceFolder 
    ) 
    foreach ($server in $servers) { 
        $serverName = $server.ServerName 
        Write-Host "Copying update files to $serverName ..." 
        # Destination folder on remote server 
        <span class="math-inline">destination \= "\\\\$serverName\\E</span>\Updates" 
        # Create destination folder if it doesn't exist 
        Invoke-Command -ComputerName $serverName -ScriptBlock { 
            param($dest) 
            if (-not (Test-Path $dest)) { 
                New-Item -Path $dest -ItemType Directory | Out-Null 
            } 
        } -ArgumentList $destination 
        # Copy all files and subfolders from source to destination 
        Copy-Item -Path (Join-Path $sourceFolder "*") -Destination $destination -Recurse -Force 
        Write-Host "Update files copied to $serverName" 
    } 
} 

# Function: Remotely install updates on selected servers (assumes updates are .msu files in D:\Updates) 
function Install-Updates-Selected { 
    Write-Host "Select the servers to install updates:" -ForegroundColor Cyan 
    for ($i = 0; $i -lt $servers.Count; $i++) { 
        Write-Host "[$i] $($servers[$i].ServerName)" 
    } 
    $selection = Read-Host "Enter the indices separated by commas (e.g., 0,2,3) or type 'all' to install on all servers" 
    if ($selection -eq "all") { 
        $selectedServers = $servers 
    } 
    else { 
        $indices = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { <span class="math-inline">\_ \-match "^\\d\+</span>" } | ForEach-Object { [int]$_ } 
        $selectedServers = @() 
        foreach ($i in $indices) { 
            if ($i -ge 0 -and $i -lt $servers.Count) { 
                $selectedServers += $servers[$i] 
            } 
            else { 
                Write-Host "Index $i is out of range. Skipping..." 
            } 
        } 
    } 
    foreach ($server in $selectedServers) { 
        $serverName = $server.ServerName 
        Write-Host "Initiating update installation on $serverName ..." 
        # Revised command using CMD for-loop: 
        # This command will iterate over MSU files in D:\Updates and run wusa.exe for each. 
        $command = 'cmd.exe /c "for %f in (D:\Updates\*.msu) do C:\Windows\System32\wusa.exe \"%f\" /quiet /norestart"' 
        # Execute the command remotely using WMI's Create method. 
        Invoke-WmiMethod -ComputerName $serverName -Class Win32_Process -Name Create -ArgumentList $command | Out-Null 
        Write-Host "Update installation command sent to $serverName" 
    } 
} 

# Function: Check update installation status with colored output and latest KB installed.
function Check-UpdatesInstallationStatus {
    foreach ($server in $servers) {
        $serverName = $server.ServerName
        Write-Host "Checking update installation status on $serverName ..." -NoNewline
        try {
            $result = Invoke-Command -ComputerName $serverName -ScriptBlock {
                # Check for pending reboot
                $pendingReboot = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
                
                # Get the latest KB installed
                $latestKB = Get-HotFix | Sort-Object -Property InstalledOn -Descending | Select-Object -First 1
                
                return @{
                    PendingReboot = $pendingReboot
                    LatestKB = $latestKB
                }
            }
            
            Write-Host ""  # New line after the "Checking..." message
            
            # Display status with appropriate color
            if ($result.PendingReboot) {
                # Green: installed and pending reboot
                Write-Host "$serverName : Installed and pending reboot." -ForegroundColor Green
            } else {
                # Yellow: no updates pending or failed
                Write-Host "$serverName : No updates pending or failed." -ForegroundColor Yellow
            }
            
            # Display latest KB information
            if ($result.LatestKB) {
                Write-Host "  Latest KB: $($result.LatestKB.HotFixID) installed on $($result.LatestKB.InstalledOn)" -ForegroundColor Cyan
                Write-Host "  Description: $($result.LatestKB.Description)" -ForegroundColor Cyan
            } else {
                Write-Host "  No hotfixes found on this server." -ForegroundColor Gray
            }
        } catch {
            Write-Host ""  # New line after the "Checking..." message
            # Red: failed to install or check status
            Write-Host "$serverName : Failed to check status. $_" -ForegroundColor Red
        }
        
        Write-Host "-----------------------------------------"
    }
}

# Function: Enable and start a specified service using WMI (for Windows Update and WinRM) 
function Enable-Start-Service { 
    param( 
        [string]$serviceName 
    ) 
    foreach ($server in $servers) { 
        $serverName = $server.ServerName 
        Write-Host "Enabling and starting service '$serviceName' on $serverName ..." 
        try { 
            $svc = Get-WmiObject -ComputerName $serverName -Class Win32_Service -Filter "Name='$serviceName'" 
            if ($svc) { 
                $svc.ChangeStartMode("Automatic") | Out-Null 
                $svc.StartService() | Out-Null 
                Write-Host "$serverName : $serviceName enabled and started." 
            } 
            else { 
                Write-Host "$serverName : Service $serviceName not found." 
            } 
        } 
        catch { 
            Write-Host "$serverName : Failed to enable/start $serviceName. $_" 
        } 
    } 
} 

# Function: Stop and disable a specified service using WMI. 
function Stop-Disable-Service { 
    param( 
        [string]$serviceName 
    ) 
    foreach ($server in $servers) { 
        $serverName = $server.ServerName 
        Write-Host "Stopping and disabling service '$serviceName' on $serverName ..." 
        try { 
            $svc = Get-WmiObject -ComputerName $serverName -Class Win32_Service -Filter "Name='$serviceName'" 
            if ($svc) { 
                $svc.StopService() | Out-Null 
                $svc.ChangeStartMode("Disabled") | Out-Null 
                Write-Host "$serverName : $serviceName stopped and disabled." 
            } 
            else { 
                Write-Host "$serverName : Service $serviceName not found." 
            } 
        } 
        catch { 
            Write-Host "$serverName : Failed to stop/disable $serviceName. $_" 
        } 
    } 
} 

# Function: Reboot the selected remote servers (instead of all at once) 
function Reboot-SelectedServers { 
    Write-Host "Select the servers to reboot:" -ForegroundColor Cyan 
    for ($i = 0; $i -lt $servers.Count; $i++) { 
        Write-Host "[$i] $($servers[$i].ServerName)" 
    } 
    $selection = Read-Host "Enter the indices separated by commas (e.g., 0,2,3) or type 'all' to reboot all servers" 
    if ($selection -eq "all") { 
        $selectedServers = $servers 
    } 
    else { 
        $indices = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { <span class="math-inline">\_ \-match "^\\d\+</span>" } | ForEach-Object { [int]$_ } 
        $selectedServers = @() 
        foreach ($i in $indices) { 
            if ($i -ge 0 -and $i -lt $servers.Count) { 
                $selectedServers += $servers[$i] 
            } 
            else { 
                Write-Host "Index $i is out of range. Skipping..." 
            } 
        } 
    } 
    foreach ($server in $selectedServers) { 
        $serverName = $server.ServerName 
        Write-Host "Rebooting $serverName ..." 
        try { 
            $os = Get-WmiObject -ComputerName $serverName -Class Win32_OperatingSystem 
            $os.Reboot() | Out-Null 
            Write-Host "$serverName : Reboot command sent." 
        } 
        catch { 
            Write-Host "$serverName : Failed to reboot. $_" 
        } 
    } 
} 

# Function: Check if the remote server is online (using ping). 
function Check-ServerStatus { 
    foreach ($server in $servers) { 
        $serverName = $server.ServerName 
        Write-Host "Checking status for $serverName ..." 
        try { 
            $online = Test-Connection -ComputerName $serverName -Count 2 -Quiet 
            if ($online) { 
                Write-Host "$serverName : Online" 
            } 
            else { 
                Write-Host "$serverName : Offline" 
            } 
        } 
        catch { 
            Write-Host "$serverName : Failed to check status. $_" 
        } 
    } 
} 

# Function: Check disk space on C: drive for all servers
function Check-CDriveSpace {
    foreach ($server in $servers) {
        $serverName = $server.ServerName
        Write-Host "Checking C: drive space on $serverName ..." -NoNewline
        
        try {
            $driveInfo = Invoke-Command -ComputerName $serverName -ScriptBlock {
                $drive = Get-PSDrive C | Select-Object Used, Free
                $totalSize = $drive.Used + $drive.Free
                $percentFree = [math]::Round(($drive.Free / $totalSize) * 100, 2)
                
                return @{
                    FreeGB = [math]::Round($drive.Free / 1GB, 2)
                    TotalGB = [math]::Round($totalSize / 1GB, 2)
                    PercentFree = $percentFree
                }
            }
            
            Write-Host ""  # New line after the "Checking..." message
            
            # Color-code based on available space
            if ($driveInfo.PercentFree -lt 10) {
                # Red for critical (less than 10% free)
                $color = "Red"
            } elseif ($driveInfo.PercentFree -lt 20) {
                # Yellow for warning (less than 20% free)
                $color = "Yellow"
            } else {
                # Green for good (20% or more free)
                $color = "Green"
            }
            
            Write-Host "$serverName : C: Drive Space" -ForegroundColor $color
            Write-Host "  Free: $($driveInfo.FreeGB) GB / $(<span class="math-inline">driveInfo\.TotalGB\) GB \(</span>($driveInfo.PercentFree)% free)" -ForegroundColor $color
            
        } catch {
            Write-Host ""  # New line after the "Checking..." message
            Write-Host "$serverName : Failed to check C: drive space. $_" -ForegroundColor Red
        }
        
        Write-Host "-----------------------------------------"
    }
}

# Function: Remove old security log archive files (older than 3 months)
function Remove-OldSecurityLogs {
    $cutoffDate = (Get-Date).AddMonths(-3)
    
    Write-Host "Select the servers to clean security logs:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $servers.Count; $i++) {
        Write-Host "[$i] $($servers[$i].ServerName)"
    }
    $selection = Read-Host "Enter the indices separated by commas (e.g., 0,2,3) or type 'all' to clean all servers"
    
    if ($selection -eq "all") {
        $selectedServers = $servers
    } else {
        $indices = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { <span class="math-inline">\_ \-match "^\\d\+</span>" } | ForEach-Object { [int]$_ }
        $selectedServers = @()
        foreach ($i in $indices) {
            if ($i -ge 0 -and $i -lt $servers.Count) {
                $selectedServers += $servers[$i]
            } else {
                Write-Host "Index $i is out of range. Skipping..." -ForegroundColor Yellow
            }
        }
    }
    
    foreach ($server in $selectedServers) {
        $serverName = $server.ServerName
        Write-Host "Cleaning old security logs on $serverName ..." -NoNewline
        
        try {
            $results = Invoke-Command -ComputerName $serverName -ScriptBlock {
                param($cutoff)
                
                $logPath = "C:\Windows\System32\winevt\Logs"
                $logFiles = Get-ChildItem -Path $logPath -Filter "Archive-Security-*" -File
                
                $removed = 0
                $totalSize = 0
                
                foreach ($file in $logFiles) {
                    if ($file.LastWriteTime -lt $cutoff) {
                        $totalSize += $file.Length
                        Remove-Item -Path $file.FullName -Force
                        $removed++
                    }
                }
                
                return @{
                    RemovedCount = $removed
                    SizeMB = [math]::Round($totalSize / 1MB, 2)
                }
            } -ArgumentList $cutoffDate
            
            Write-Host ""  # New line after the "Cleaning..." message
            
            if ($results.RemovedCount -gt 0) {
                Write-Host "$serverName : Removed $($results.RemovedCount) security log files older than 3 months" -ForegroundColor Green
                Write-Host "  Freed approximately $($results.SizeMB) MB of disk space" -ForegroundColor Green
            } else {
                Write-Host "$serverName : No old security log files found to remove" -ForegroundColor Yellow
            }
            
        } catch {
            Write-Host ""  # New line after the "Cleaning..." message
            Write-Host "$serverName : Failed to clean security logs. $_" -ForegroundColor Red
        }
        
        Write-Host "-----------------------------------------"
    }
}

# Function: Check Edge version on selected servers
function Check-EdgeVersion {
    Write-Host "Select the servers to check Edge version:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $servers.Count; $i++) {
        Write-Host "[$i] $($servers[$i].ServerName)"
    }
    $selection = Read-Host "Enter the indices separated by commas (e.g., 0,2,3) or type 'all' to check all servers"
    
    if ($selection -eq "all") {
        $selectedServers = $servers
    } else {
        $indices = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { <span class="math-inline">\_ \-match "^\\d\+</span>" } | ForEach-Object { [int]$_ }
        $selectedServers = @()
        foreach ($i in $indices) {
            if ($i -ge 0 -and $i -lt $servers.Count) {
                $selectedServers += $servers[$i]
            } else {
                Write-Host "Index $i is out of range. Skipping..." -ForegroundColor Yellow
            }
        }
    }
    
    foreach ($server in $selectedServers) {
        $serverName = $server.ServerName
        Write-Host "Checking Microsoft Edge version on $serverName ..." -NoNewline
        
        try {
            $edgeInfo = Invoke-Command -ComputerName $serverName -ScriptBlock {
                $edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application"
                if (Test-Path $edgePath) {
                    # Try to get version from version.dll first
                    if (Test-Path "$edgePath\version") {
                        $versionFile = Get-Item "$edgePath\version"
                        $version = $versionFile.Name
                        $installDate = $versionFile.CreationTime
                        return @{
                            Installed = $true
                            Version = $version
                            InstallDate = $installDate
                            Path = $edgePath
                        }
                    } else {
                        # Get first directory that looks like a version number
                        $versionFolders = Get-ChildItem -Path $edgePath -Directory | 
                                         Where-Object { <span class="math-inline">\_\.Name \-match '^\\d\+\\\.\\d\+\\\.\\d\+\\\.\\d\+</span>' } | 
                                         Sort-Object Name -Descending
                        if ($versionFolders.Count -gt 0) {
                            $version = $versionFolders[0].Name
                            $installDate = $versionFolders[0].CreationTime
                            return @{
                                Installed = $true
                                Version = $version
                                InstallDate = $installDate
                                Path = $edgePath
                            }
                        } else {
                            return @{
                                Installed = $true
                                Version = "Unknown"
                                InstallDate = $null
                                Path = $edgePath
                            }
                        }
                    }
                } else {
                    # Check alternative location
                    $edgePath = "C:\Program Files\Microsoft\Edge\Application"
                    if (Test-Path $edgePath) {
                        # Try to get version from version.dll first
                        if (Test-Path "$edgePath\version") {
                            $versionFile = Get-Item "$edgePath\version"
                            $version = $versionFile.Name
                            $installDate = $versionFile.CreationTime
                            return @{
                                Installed = $true
                                Version = $version
                                InstallDate = $installDate
                                Path = $edgePath
                            }
                        } else {
                            $versionFolders = Get-ChildItem -Path $edgePath -Directory | 
                                             Where-Object { <span class="math-inline">\_\.Name \-match '^\\d\+\\\.\\d\+\\\.\\d\+\\\.\\d\+</span>' } | 
                                             Sort-Object Name -Descending
                            if ($versionFolders.Count -gt 0) {
                                $version = $versionFolders[0].Name
                                $installDate = $versionFolders[0].CreationTime
                                return @{
                                    Installed = $true
                                    Version = $version
                                    InstallDate = $installDate
                                    Path = $edgePath
                                }
                            } else {
                                return @{
                                    Installed = $true
                                    Version = "Unknown"
                                    InstallDate = $null
                                    Path = $edgePath
                                }
                            }
                        }
                    } else {
                        return @{
                            Installed = $false
                            Version = "Not Installed"
                            InstallDate = $null
                            Path = $null
                        }
                    }
                }
            }
            
            Write-Host ""  # New line after the "Checking..." message
            
            if ($edgeInfo.Installed) {
                Write-Host "$serverName : Microsoft Edge" -ForegroundColor Cyan
                Write-Host "  Version: $($edgeInfo.Version)" -ForegroundColor Cyan
                if ($edgeInfo.InstallDate) {
                    Write-Host "  Install Date: $($edgeInfo.InstallDate)" -ForegroundColor Cyan
                }
                Write-Host "  Path: $($edgeInfo.Path)" -ForegroundColor Cyan
            } else {
                Write-Host "$serverName : Microsoft Edge is not installed" -ForegroundColor Yellow
            }
            
        } catch {
            Write-Host ""  # New line after the "Checking..." message
            Write-Host "$serverName : Failed to check Edge version. $_" -ForegroundColor Red
        }
        
        Write-Host "-----------------------------------------"
    }
}

# Function: Update Edge on selected servers
function Update-Edge {
    Write-Host "Select the servers to update Microsoft Edge:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $servers.Count; $i++) {
        Write-Host "[$i] $($servers[$i].ServerName)"
    }
    $selection = Read-Host "Enter the indices separated by commas (e.g., 0,2,3) or type 'all' to update all servers"
    
    if ($selection -eq "all") {
        $selectedServers = $servers
    } else {
        $indices = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" } | ForEach-Object { [int]$_ }
        $selectedServers = @()
        foreach ($i in $indices) {
            if ($i -ge 0 -and $i -lt $servers.Count) {
                $selectedServers += $servers[$i]
            } else {
                Write-Host "Index $i is out of range. Skipping..." -ForegroundColor Yellow
            }
        }
    }
    
    # Look for Edge MSI files in the D:\Updates folder on each server
    foreach ($server in $selectedServers) {
        $serverName = $server.ServerName
        Write-Host "Checking for Edge MSI on $serverName ..." -NoNewline
        
        try {
            $updateResult = Invoke-Command -ComputerName $serverName -ScriptBlock {
                $updateFolder = "D:\Updates"
                
                # Check if the updates folder exists
                if (-not (Test-Path $updateFolder)) {
                    return @{
                        Status = "Error"
                        Message = "Updates folder not found at D:\Updates"
                    }
                }
                
                # Find Edge MSI files
                $edgeMsiFiles = Get-ChildItem -Path $updateFolder -Filter "*edge*.msi" -File
                
                if ($edgeMsiFiles.Count -eq 0) {
                    return @{
                        Status = "Error"
                        Message = "No Edge MSI files found in D:\Updates"
                    }
                }
                
                # Find the newest Edge MSI file
                $newestEdgeMsi = $edgeMsiFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                
                # Install Edge using the MSI
                try {
                    $installProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$($newestEdgeMsi.FullName)`" /qn /norestart" -Wait -PassThru -NoNewWindow
                    
                    if ($installProcess.ExitCode -eq 0) {
                        return @{
                            Status = "Success"
                            Message = "Edge successfully updated using $($newestEdgeMsi.Name)"
                            ExitCode = $installProcess.ExitCode
                        }
                    } else {
                        return @{
                            Status = "Error"
                            Message = "Edge installation failed"
                            MsiFile = $newestEdgeMsi.Name
                            ExitCode = $installProcess.ExitCode
                        }
                    }
                } catch {
                    return @{
                        Status = "Error"
                        Message = "Error executing MSI: $_"
                        MsiFile = $newestEdgeMsi.Name
                    }
                }
            }
            
            Write-Host ""  # New line after the "Checking..." message
            
            if ($updateResult.Status -eq "Success") {
                Write-Host "$serverName : $($updateResult.Message)" -ForegroundColor Green
            } else {
                Write-Host "$serverName : $($updateResult.Message)" -ForegroundColor Red
                if ($updateResult.ContainsKey("MsiFile")) {
                    Write-Host "  MSI File: $($updateResult.MsiFile)" -ForegroundColor Yellow
                }
                if ($updateResult.ContainsKey("ExitCode")) {
                    Write-Host "  Exit Code: $($updateResult.ExitCode)" -ForegroundColor Yellow
                }
            }
            
        } catch {
            Write-Host ""  # New line after the "Checking..." message
            Write-Host "$serverName : Failed to update Edge. $_" -ForegroundColor Red
        }
        
        Write-Host "-----------------------------------------"
    }
}

# --- Main Menu Loop ---
do {
    Write-Host ""
    Write-Host "==================== MENU ====================" -ForegroundColor Green
    Write-Host "1. Copy Updates Files"
    Write-Host "2. Install Updates (Selected Servers)"
    Write-Host "3. Check Updates Installation Status"
    Write-Host "4. Enable and Start Windows Update Service"
    Write-Host "5. Stop and Disable Windows Update Service"
    Write-Host "6. Enable and Start WinRM Service (WMI)"
    Write-Host "7. Stop and Disable WinRM Service (WMI)"
    Write-Host "8. Reboot Selected Servers"
    Write-Host "9. Check Server Status"
    Write-Host "10. Check C: Drive Space"
    Write-Host "11. Clean Old Security Logs (3+ months)"
    Write-Host "12. Check Edge Version"
    Write-Host "13. Update Edge"
    Write-Host "14. Quit"
    Write-Host "==============================================" -ForegroundColor Green
    $choice = Read-Host "Enter your choice (1-14)"
    
    switch ($choice) {
        "1" {
            $sourceFolder = Read-Host "Enter the folder path containing the update files (e.g., D:\Updates)"
            Copy-UpdatesFiles -sourceFolder $sourceFolder
        }
        "2" { Install-Updates-Selected }
        "3" { Check-UpdatesInstallationStatus }
        "4" { Enable-Start-Service -serviceName "wuauserv" }
        "5" { Stop-Disable-Service -serviceName "wuauserv" }
        "6" { Enable-Start-Service -serviceName "WinRM" }
        "7" { Stop-Disable-Service -serviceName "WinRM" }
        "8" { Reboot-SelectedServers }
        "9" { Check-ServerStatus }
        "10" { Check-CDriveSpace }
        "11" { Remove-OldSecurityLogs }
        "12" { Check-EdgeVersion }
        "13" { Update-Edge }
        "14" {
            Write-Host "Exiting..."
        }
        default {
            Write-Host "Invalid option. Please try again." -ForegroundColor Yellow
        }
    }

    # Pause for the user to read the output before showing the menu again
    if ($choice -ne "14") {
        Read-Host "Press Enter to return to the menu..."
    }

} while ($choice -ne "14") 
# --- END SCRIPT ---
