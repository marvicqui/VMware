#Requires -Version 5.1
#Requires -Modules VMware.PowerCLI

# --- Main Menu ---
function Show-MainMenu {
    Write-Host "=================================="
    Write-Host "System Management Menu"
    Write-Host "=================================="
    Write-Host "1. Snapshot Management (vCenter)"
    Write-Host "2. Server Management (Windows)"
    Write-Host "3. VM Management (vCenter)"
    Write-Host "4. Quit"
    Write-Host "=================================="
    $choice = Read-Host "Select an option (1-4)"
    return $choice
}

# --- Snapshot Management Submenu ---
function Show-SnapshotMenu {
    Write-Host "=================================="
    Write-Host "Snapshot Management Menu"
    Write-Host "=================================="
    Write-Host "1. Take Snapshots"
    Write-Host "2. Delete Last Snapshot"
    Write-Host "3. Delete All Snapshots"
    Write-Host "4. Delete Selected Snapshots"
    Write-Host "5. Restart Guest OS"
    Write-Host "6. Back to Main Menu"
    Write-Host "=================================="
    $choice = Read-Host "Select an option (1-6)"
    return $choice
}

# --- Server Management Submenu ---
function Show-ServerMenu {
    Write-Host "=================================="
    Write-Host "Server Management Menu"
    Write-Host "=================================="
    Write-Host "1. Copy Update Files"
    Write-Host "2. Install Updates"
    Write-Host "3. Check Update Status"
    Write-Host "4. Enable and Start Windows Update Service"
    Write-Host "5. Stop and Disable Windows Update Service"
    Write-Host "6. Enable and Start WinRM Service"
    Write-Host "7. Stop and Disable WinRM Service"
    Write-Host "8. Reboot Servers"
    Write-Host "9. Check Server Status"
    Write-Host "10. Check C: Drive Space"
    Write-Host "11. Clean Old Security Logs"
    Write-Host "12. Check Edge Version"
    Write-Host "13. Update Edge"
    Write-Host "14. Back to Main Menu"
    Write-Host "=================================="
    $choice = Read-Host "Select an option (1-14)"
    return $choice
}

# --- VM Management Submenu ---
function Show-VMMenu {
    Write-Host "=================================="
    Write-Host "VM Management Menu"
    Write-Host "=================================="
    Write-Host "1. Connect to vCenter"
    Write-Host "2. Clone VM"
    Write-Host "3. Disconnect vNIC"
    Write-Host "4. Connect vNIC"
    Write-Host "5. Rename Computer"
    Write-Host "6. Set Password for CtxAdmin2"
    Write-Host "7. Start VM"
    Write-Host "8. Extend Partition"
    Write-Host "9. Back to Main Menu"
    Write-Host "=================================="
    $choice = Read-Host "Select an option (1-9)"
    return $choice
}

# --- Snapshot Management Functions ---
$vCenters = @("vwvsinap01.edificios.gfbanorte", "vmwvvcinapha01.edificios.gfbanorte", "vmwvcsttapha01.edificios.gfbanorte")
$logFilePath = "$env:USERPROFILE\Desktop\SnapshotManagementLog.txt"

function Log-Message {
    param ([Parameter(Mandatory=$true)][string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Add-Content -Path $logFilePath -Value $logEntry
    Write-Host $logEntry
}

function Select-vCenter {
    while ($true) {
        Write-Host "Select a vCenter to connect to:"
        for ($i = 0; $i -lt $vCenters.Length; $i++) {
            Write-Host "$($i + 1). $($vCenters[$i])"
        }
        $selection = Read-Host "Enter the number of the vCenter"
        if ($selection -ge 1 -and $selection -le $vCenters.Length) {
            return $vCenters[$selection - 1]
        } else {
            Write-Host "Invalid selection. Please enter a number between 1 and $($vCenters.Length)."
        }
    }
}

function Take-VMSnapshot {
    param ([Parameter(Mandatory=$true)][string[]]$VMNames, [Parameter(Mandatory=$true)][string]$SnapshotName, [Parameter(Mandatory=$true)][string]$SnapshotDescription)
    foreach ($vmName in $VMNames) {
        try {
            $vm = Get-VM -Name $vmName -ErrorAction Stop
            New-Snapshot -VM $vm -Name $SnapshotName -Description $SnapshotDescription -Confirm:$false -ErrorAction Stop
            Log-Message "Snapshot taken for VM: $vmName"
        } catch {
            Log-Message "Failed to take snapshot for VM: $vmName. Error: $_"
        }
    }
}

function Remove-LastSnapshot {
    param ([Parameter(Mandatory=$true)][string[]]$VMNames)
    foreach ($vmName in $VMNames) {
        try {
            $vm = Get-VM -Name $vmName -ErrorAction Stop
            $lastSnapshot = Get-Snapshot -VM $vm | Sort-Object Created -Descending | Select-Object -First 1 -ErrorAction Stop
            if ($lastSnapshot) {
                Remove-Snapshot -Snapshot $lastSnapshot -Confirm:$false -ErrorAction Stop
                Log-Message "Last snapshot deleted for VM: $vmName"
            } else {
                Log-Message "No snapshots found for VM: $vmName"
            }
        } catch {
            Log-Message "Failed to delete last snapshot for VM: $vmName. Error: $_"
        }
    }
}

function Remove-AllSnapshots {
    param ([Parameter(Mandatory=$true)][string[]]$VMNames)
    foreach ($vmName in $VMNames) {
        try {
            $vm = Get-VM -Name $vmName -ErrorAction Stop
            $snapshots = Get-Snapshot -VM $vm -ErrorAction Stop
            if ($snapshots) {
                Remove-Snapshot -Snapshot $snapshots -Confirm:$false -ErrorAction Stop
                Log-Message "All snapshots deleted for VM: $vmName"
            } else {
                Log-Message "No snapshots found for VM: $vmName"
            }
        } catch {
            Log-Message "Failed to delete all snapshots for VM: $vmName. Error: $_"
        }
    }
}

function Remove-SelectedSnapshots {
    param ([Parameter(Mandatory=$true)][string[]]$VMNames)
    foreach ($vmName in $VMNames) {
        try {
            $vm = Get-VM -Name $vmName -ErrorAction Stop
            $snapshots = Get-Snapshot -VM $vm -ErrorAction Stop
            if ($snapshots) {
                $selectedSnapshots = $snapshots | Out-GridView -Title "Select Snapshots to Delete for VM: $vmName" -PassThru
                if ($selectedSnapshots) {
                    Remove-Snapshot -Snapshot $selectedSnapshots -Confirm:$false -ErrorAction Stop
                    Log-Message "Selected snapshots deleted for VM: $vmName"
                } else {
                    Log-Message "No snapshots selected for VM: $vmName"
                }
            } else {
                Log-Message "No snapshots found for VM: $vmName"
            }
        } catch {
            Log-Message "Failed to delete selected snapshots for VM: $vmName. Error: $_"
        }
    }
}

function Restart-GuestOS {
    param ([Parameter(Mandatory=$true)][string[]]$VMNames)
    foreach ($vmName in $VMNames) {
        try {
            $vm = Get-VM -Name $vmName -ErrorAction Stop
            if ($vm.ExtensionData.Guest.ToolsRunningStatus -eq "guestToolsRunning") {
                Restart-VMGuest -VM $vm -Confirm:$false -ErrorAction Stop
                Log-Message "Guest OS restart initiated for VM: $vmName"
            } else {
                Log-Message "VMware Tools is not running on VM: $vmName. Cannot restart guest OS."
            }
        } catch {
            Log-Message "Failed to restart guest OS for VM: $vmName. Error: $_"
        }
    }
}

function Get-CSVFilePath {
    Add-Type -AssemblyName System.Windows.Forms
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $fileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $fileDialog.Title = "Select a CSV file with VM names"
    if ($fileDialog.ShowDialog() -eq 'OK') {
        return $fileDialog.FileName
    } else {
        Log-Message "No file selected. Exiting script."
        Disconnect-VIServer -Confirm:$false
        exit
    }
}

function Select-VMsFromList {
    $vms = Get-VM | Sort-Object Name
    $selectedVMs = $vms | Out-GridView -Title "Select VMs" -PassThru | Select-Object -ExpandProperty Name
    return $selectedVMs
}

# --- Server Management Functions ---
function Copy-UpdatesFiles {
    param ([string]$sourceFolder)
    foreach ($server in $servers) {
        $serverName = $server.ServerName
        Write-Host "Copying update files to $serverName ..."
        $destination = "\\$serverName\E$\Updates"
        Invoke-Command -ComputerName $serverName -ScriptBlock { param($dest) if (-not (Test-Path $dest)) { New-Item -Path $dest -ItemType Directory | Out-Null } } -ArgumentList $destination
        Copy-Item -Path (Join-Path $sourceFolder "*") -Destination $destination -Recurse -Force
        Write-Host "Update files copied to $serverName"
    }
}

function Install-Updates-Selected {
    Write-Host "Select the servers to install updates:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $servers.Count; $i++) { Write-Host "[$i] $($servers[$i].ServerName)" }
    $selection = Read-Host "Enter the indices separated by commas (e.g., 0,2,3) or type 'all' to install on all servers"
    if ($selection -eq "all") { $selectedServers = $servers }
    else {
        $indices = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" } | ForEach-Object { [int]$_ }
        $selectedServers = @()
        foreach ($i in $indices) {
            if ($i -ge 0 -and $i -lt $servers.Count) { $selectedServers += $servers[$i] }
            else { Write-Host "Index $i is out of range. Skipping..." }
        }
    }
    foreach ($server in $selectedServers) {
        $serverName = $server.ServerName
        Write-Host "Initiating update installation on $serverName ..."
        $command = 'cmd.exe /c "for %f in (D:\Updates\*.msu) do C:\Windows\System32\wusa.exe \"%f\" /quiet /norestart"'
        Invoke-WmiMethod -ComputerName $serverName -Class Win32_Process -Name Create -ArgumentList $command | Out-Null
        Write-Host "Update installation command sent to $serverName"
    }
}

function Check-UpdatesInstallationStatus {
    foreach ($server in $servers) {
        $serverName = $server.ServerName
        Write-Host "Checking update installation status on $serverName ..." -NoNewline
        try {
            $result = Invoke-Command -ComputerName $serverName -ScriptBlock {
                $pendingReboot = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
                $latestKB = Get-HotFix | Sort-Object -Property InstalledOn -Descending | Select-Object -First 1
                return @{ PendingReboot = $pendingReboot; LatestKB = $latestKB }
            }
            Write-Host ""
            if ($result.PendingReboot) { Write-Host "$serverName : Installed and pending reboot." -ForegroundColor Green }
            else { Write-Host "$serverName : No updates pending or failed." -ForegroundColor Yellow }
            if ($result.LatestKB) {
                Write-Host "  Latest KB: $($result.LatestKB.HotFixID) installed on $($result.LatestKB.InstalledOn)" -ForegroundColor Cyan
                Write-Host "  Description: $($result.LatestKB.Description)" -ForegroundColor Cyan
            } else { Write-Host "  No hotfixes found on this server." -ForegroundColor Gray }
        } catch {
            Write-Host ""
            Write-Host "$serverName : Failed to check status. $_" -ForegroundColor Red
        }
        Write-Host "-----------------------------------------"
    }
}

function Enable-Start-Service {
    param ([string]$serviceName)
    foreach ($server in $servers) {
        $serverName = $server.ServerName
        Write-Host "Enabling and starting service '$serviceName' on $serverName ..."
        try {
            $svc = Get-WmiObject -ComputerName $serverName -Class Win32_Service -Filter "Name='$serviceName'"
            if ($svc) {
                $svc.ChangeStartMode("Automatic") | Out-Null
                $svc.StartService() | Out-Null
                Write-Host "$serverName : $serviceName enabled and started."
            } else { Write-Host "$serverName : Service $serviceName not found." }
        } catch { Write-Host "$serverName : Failed to enable/start $serviceName. $_" }
    }
}

function Stop-Disable-Service {
    param ([string]$serviceName)
    foreach ($server in $servers) {
        $serverName = $server.ServerName
        Write-Host "Stopping and disabling service '$serviceName' on $serverName ..."
        try {
            $svc = Get-WmiObject -ComputerName $serverName -Class Win32_Service -Filter "Name='$serviceName'"
            if ($svc) {
                $svc.StopService() | Out-Null
                $svc.ChangeStartMode("Disabled") | Out-Null
                Write-Host "$serverName : $serviceName stopped and disabled."
            } else { Write-Host "$serverName : Service $serviceName not found." }
        } catch { Write-Host "$serverName : Failed to stop/disable $serviceName. $_" }
    }
}

function Reboot-SelectedServers {
    Write-Host "Select the servers to reboot:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $servers.Count; $i++) { Write-Host "[$i] $($servers[$i].ServerName)" }
    $selection = Read-Host "Enter the indices separated by commas (e.g., 0,2,3) or type 'all' to reboot all servers"
    if ($selection -eq "all") { $selectedServers = $servers }
    else {
        $indices = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" } | ForEach-Object { [int]$_ }
        $selectedServers = @()
        foreach ($i in $indices) {
            if ($i -ge 0 -and $i -lt $servers.Count) { $selectedServers += $servers[$i] }
            else { Write-Host "Index $i is out of range. Skipping..." }
        }
    }
    foreach ($server in $selectedServers) {
        $serverName = $server.ServerName
        Write-Host "Rebooting $serverName ..."
        try {
            $os = Get-WmiObject -ComputerName $serverName -Class Win32_OperatingSystem
            $os.Reboot() | Out-Null
            Write-Host "$serverName : Reboot command sent."
        } catch { Write-Host "$serverName : Failed to reboot. $_" }
    }
}

function Check-ServerStatus {
    foreach ($server in $servers) {
        $serverName = $server.ServerName
        Write-Host "Checking status for $serverName ..."
        try {
            $online = Test-Connection -ComputerName $serverName -Count 2 -Quiet
            if ($online) { Write-Host "$serverName : Online" } else { Write-Host "$serverName : Offline" }
        } catch { Write-Host "$serverName : Failed to check status. $_" }
    }
}

function Check-CDriveSpace {
    foreach ($server in $servers) {
        $serverName = $server.ServerName
        Write-Host "Checking C: drive space on $serverName ..." -NoNewline
        try {
            $driveInfo = Invoke-Command -ComputerName $serverName -ScriptBlock {
                $drive = Get-PSDrive C | Select-Object Used, Free
                $totalSize = $drive.Used + $drive.Free
                $percentFree = [math]::Round(($drive.Free / $totalSize) * 100, 2)
                return @{ FreeGB = [math]::Round($drive.Free / 1GB, 2); TotalGB = [math]::Round($totalSize / 1GB, 2); PercentFree = $percentFree }
            }
            Write-Host ""
            $color = if ($driveInfo.PercentFree -lt 10) { "Red" } elseif ($driveInfo.PercentFree -lt 20) { "Yellow" } else { "Green" }
            Write-Host "$serverName : C: Drive Space" -ForegroundColor $color
            Write-Host "  Free: $($driveInfo.FreeGB) GB / $($driveInfo.TotalGB) GB ($($driveInfo.PercentFree)% free)" -ForegroundColor $color
        } catch {
            Write-Host ""
            Write-Host "$serverName : Failed to check C: drive space. $_" -ForegroundColor Red
        }
        Write-Host "-----------------------------------------"
    }
}

function Remove-OldSecurityLogs {
    $cutoffDate = (Get-Date).AddMonths(-3)
    Write-Host "Select the servers to clean security logs:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $servers.Count; $i++) { Write-Host "[$i] $($servers[$i].ServerName)" }
    $selection = Read-Host "Enter the indices separated by commas (e.g., 0,2,3) or type 'all' to clean all servers"
    if ($selection -eq "all") { $selectedServers = $servers }
    else {
        $indices = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" } | ForEach-Object { [int]$_ }
        $selectedServers = @()
        foreach ($i in $indices) {
            if ($i -ge 0 -and $i -lt $servers.Count) { $selectedServers += $servers[$i] }
            else { Write-Host "Index $i is out of range. Skipping..." -ForegroundColor Yellow }
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
                return @{ RemovedCount = $removed; SizeMB = [math]::Round($totalSize / 1MB, 2) }
            } -ArgumentList $cutoffDate
            Write-Host ""
            if ($results.RemovedCount -gt 0) {
                Write-Host "$serverName : Removed $($results.RemovedCount) security log files older than 3 months" -ForegroundColor Green
                Write-Host "  Freed approximately $($results.SizeMB) MB of disk space" -ForegroundColor Green
            } else { Write-Host "$serverName : No old security log files found to remove" -ForegroundColor Yellow }
        } catch {
            Write-Host ""
            Write-Host "$serverName : Failed to clean security logs. $_" -ForegroundColor Red
        }
        Write-Host "-----------------------------------------"
    }
}

function Check-EdgeVersion {
    Write-Host "Select the servers to check Edge version:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $servers.Count; $i++) { Write-Host "[$i] $($servers[$i].ServerName)" }
    $selection = Read-Host "Enter the indices separated by commas (e.g., 0,2,3) or type 'all' to check all servers"
    if ($selection -eq "all") { $selectedServers = $servers }
    else {
        $indices = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" } | ForEach-Object { [int]$_ }
        $selectedServers = @()
        foreach ($i in $indices) {
            if ($i -ge 0 -and $i -lt $servers.Count) { $selectedServers += $servers[$i] }
            else { Write-Host "Index $i is out of range. Skipping..." -ForegroundColor Yellow }
        }
    }
    foreach ($server in $selectedServers) {
        $serverName = $server.ServerName
        Write-Host "Checking Microsoft Edge version on $serverName ..." -NoNewline
        try {
            $edgeInfo = Invoke-Command -ComputerName $serverName -ScriptBlock {
                $edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application"
                if (Test-Path $edgePath) {
                    $versionFolders = Get-ChildItem -Path $edgePath -Directory | Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } | Sort-Object Name -Descending
                    if ($versionFolders.Count -gt 0) {
                        return @{ Installed = $true; Version = $versionFolders[0].Name; InstallDate = $versionFolders[0].CreationTime }
                    } else { return @{ Installed = $true; Version = "Unknown"; InstallDate = $null } }
                } else { return @{ Installed = $false; Version = "Not Installed"; InstallDate = $null } }
            }
            Write-Host ""
            if ($edgeInfo.Installed) {
                Write-Host "$serverName : Microsoft Edge" -ForegroundColor Cyan
                Write-Host "  Version: $($edgeInfo.Version)" -ForegroundColor Cyan
                if ($edgeInfo.InstallDate) { Write-Host "  Install Date: $($edgeInfo.InstallDate)" -ForegroundColor Cyan }
            } else { Write-Host "$serverName : Microsoft Edge is not installed" -ForegroundColor Yellow }
        } catch {
            Write-Host ""
            Write-Host "$serverName : Failed to check Edge version. $_" -ForegroundColor Red
        }
        Write-Host "-----------------------------------------"
    }
}

function Update-Edge {
    Write-Host "Select the servers to update Microsoft Edge:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $servers.Count; $i++) { Write-Host "[$i] $($servers[$i].ServerName)" }
    $selection = Read-Host "Enter the indices separated by commas (e.g., 0,2,3) or type 'all' to update all servers"
    if ($selection -eq "all") { $selectedServers = $servers }
    else {
        $indices = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match "^\d+$" } | ForEach-Object { [int]$_ }
        $selectedServers = @()
        foreach ($i in $indices) {
            if ($i -ge 0 -and $i -lt $servers.Count) { $selectedServers += $servers[$i] }
            else { Write-Host "Index $i is out of range. Skipping..." -ForegroundColor Yellow }
        }
    }
    foreach ($server in $selectedServers) {
        $serverName = $server.ServerName
        Write-Host "Checking for Edge MSI on $serverName ..." -NoNewline
        try {
            $updateResult = Invoke-Command -ComputerName $serverName -ScriptBlock {
                $updateFolder = "D:\Updates"
                if (-not (Test-Path $updateFolder)) { return @{ Status = "Error"; Message = "Updates folder not found at D:\Updates" } }
                $edgeMsiFiles = Get-ChildItem -Path $updateFolder -Filter "*edge*.msi" -File
                if ($edgeMsiFiles.Count -eq 0) { return @{ Status = "Error"; Message = "No Edge MSI files found in D:\Updates" } }
                $newestEdgeMsi = $edgeMsiFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                $installProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$($newestEdgeMsi.FullName)`" /qn /norestart" -Wait -PassThru -NoNewWindow
                if ($installProcess.ExitCode -eq 0) { return @{ Status = "Success"; Message = "Edge successfully updated using $($newestEdgeMsi.Name)" } }
                else { return @{ Status = "Error"; Message = "Edge installation failed"; MsiFile = $newestEdgeMsi.Name; ExitCode = $installProcess.ExitCode } }
            }
            Write-Host ""
            if ($updateResult.Status -eq "Success") { Write-Host "$serverName : $($updateResult.Message)" -ForegroundColor Green }
            else {
                Write-Host "$serverName : $($updateResult.Message)" -ForegroundColor Red
                if ($updateResult.ContainsKey("MsiFile")) { Write-Host "  MSI File: $($updateResult.MsiFile)" -ForegroundColor Yellow }
                if ($updateResult.ContainsKey("ExitCode")) { Write-Host "  Exit Code: $($updateResult.ExitCode)" -ForegroundColor Yellow }
            }
        } catch {
            Write-Host ""
            Write-Host "$serverName : Failed to update Edge. $_" -ForegroundColor Red
        }
        Write-Host "-----------------------------------------"
    }
}

# --- VM Management Functions ---
function Clone-VM {
    if (-not $global:DefaultVIServer) { Write-Host "Error: Not connected to a vCenter Server. Please connect first."; return }
    Write-Host "Cloning VM..."
    $sourceVMName = Read-Host "Please enter the name of the source VM to clone"
    $newVMName = Read-Host "Please enter the name for the new VM"
    try {
        $sourceVM = Get-VM -Name $sourceVMName -ErrorAction Stop
        $newVM = New-VM -VM $sourceVM -Name $newVMName -VMHost $sourceVM.VMHost -Datastore "VMsRecursosTercerosHA" -ErrorAction Stop
        Write-Host "VM '$newVMName' has been cloned successfully from '$sourceVMName'."
    } catch { Write-Host "Error cloning VM: $_" }
}

function Disconnect-vNIC {
    if (-not $global:DefaultVIServer) { Write-Host "Error: Not connected to a vCenter Server. Please connect first."; return }
    Write-Host "Disconnecting vNIC..."
    $vmName = Read-Host "Please enter the name of the VM to configure"
    try {
        $vm = Get-VM -Name $vmName -ErrorAction Stop
        Get-NetworkAdapter -VM $vm | Set-NetworkAdapter -NetworkName "VDI_Terceros_2067" -StartConnected:$false -Confirm:$false -ErrorAction Stop
        Write-Host "VM '$vmName' has been configured with VLAN 'VDI_Terceros_2067' and disconnected network adapter."
    } catch { Write-Host "Error disconnecting vNIC: $_" }
}

function Connect-vNIC {
    if (-not $global:DefaultVIServer) { Write-Host "Error: Not connected to a vCenter Server. Please connect first."; return }
    Write-Host "Connecting vNIC..."
    $vmName = Read-Host "Please enter the name of the VM to configure"
    try {
        $vm = Get-VM -Name $vmName -ErrorAction Stop
        Get-NetworkAdapter -VM $vm | Set-NetworkAdapter -Connected:$true -Confirm:$false -ErrorAction Stop
        $updatedAdapter = Get-NetworkAdapter -VM $vm
        if ($updatedAdapter.ConnectionState.Connected) { Write-Host "Success: vNIC for VM '$vmName' is now connected to VLAN '$($updatedAdapter.NetworkName)'." }
        else { Write-Host "Error: vNIC for VM '$vmName' is still disconnected." }
    } catch { Write-Host "Error connecting vNIC: $_" }
}

function Rename-Computer {
    if (-not $global:DefaultVIServer) { Write-Host "Error: Not connected to a vCenter Server. Please connect first."; return }
    Write-Host "Renaming computer..."
    $vmName = Read-Host "Please enter the name of the VM to configure"
    try {
        $vm = Get-VM -Name $vmName -ErrorAction Stop
        $username = "$vmName\ctxadmin"
        $password = "Banorte2020."
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        $guestCredential = New-Object System.Management.Automation.PSCredential($username, $securePassword)
        if ($vm.PowerState -ne "PoweredOn") {
            Write-Host "Powering on VM '$vmName'..."
            Start-VM -VM $vm -Confirm:$false
            Start-Sleep -Seconds 60
        }
        $newComputerName = Read-Host "Please enter the new computer name for VM '$vmName'"
        $workgroupName = Read-Host "Please enter the workgroup name for the VM (e.g., WORKGROUP)"
        $guestScript = @"
        Remove-Computer -Force -PassThru
        Add-Computer -WorkgroupName "$workgroupName" -Force
        Rename-Computer -NewName "$newComputerName" -Force
        Write-Output "Computer name set to $newComputerName and joined workgroup $workgroupName. Rebooting now..."
        Restart-Computer -Force
"@
        Write-Host "Applying computer name and workgroup changes to VM '$vmName'..."
        $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
        Write-Host "Script output: $($result.ScriptOutput)"
        Write-Host "VM '$vmName' has been configured with the new computer name '$newComputerName' and workgroup '$workgroupName'. The VM is rebooting to apply changes."
        Write-Host "Note: The computer object for '$vmName' may still exist in Active Directory. Please manually remove it from the domain to avoid stale entries."
    } catch { Write-Host "Error renaming computer: $_" }
}

function Set-CtxAdmin2Password {
    if (-not $global:DefaultVIServer) { Write-Host "Error: Not connected to a vCenter Server. Please connect first."; return }
    Write-Host "Setting password for ctxadmin2..."
    $vmName = Read-Host "Please enter the name of the VM to configure"
    try {
        $vm = Get-VM -Name $vmName -ErrorAction Stop
        $username = "$vmName\ctxadmin"
        $password = "Banorte2020."
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        $guestCredential = New-Object System.Management.Automation.PSCredential($username, $securePassword)
        if ($vm.PowerState -ne "PoweredOn") {
            Write-Host "Powering on VM '$vmName'..."
            Start-VM -VM $vm -Confirm:$false
            Start-Sleep -Seconds 60
        }
        $newPassword = "VNTB4n0rt3B4n0rt3"
        $guestScript = @"
        try {
            `$user = [ADSI]('WinNT://./ctxadmin2,user')
            if (`$user -eq `$null) { Write-Output 'Error: User ctxadmin2 not found.'; exit 1 }
            `$user.SetPassword('$newPassword')
            `$user.SetInfo()
            Write-Output 'Password for ctxadmin2 changed successfully.'
        } catch { Write-Output 'Error changing password for ctxadmin2: `$_'; exit 1 }
"@
        Write-Host "Changing password for ctxadmin2 on VM '$vmName'..."
        $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
        Write-Host "Script output: $($result.ScriptOutput)"
        Write-Host "Password change completed for user ctxadmin2 on VM '$vmName'."
    } catch { Write-Host "Error setting password for ctxadmin2: $_" }
}

function Start-VMInstance {
    if (-not $global:DefaultVIServer) { Write-Host "Error: Not connected to a vCenter Server. Please connect first."; return }
    Write-Host "Starting VM..."
    $vmName = Read-Host "Please enter the name of the VM to start"
    try {
        Start-VM -VM $vmName -Confirm:$false -ErrorAction Stop
        Write-Host "VM '$vmName' has been started."
    } catch { Write-Host "Error starting VM: $_" }
}

function Extend-Partition {
    if (-not $global:DefaultVIServer) { Write-Host "Error: Not connected to a vCenter Server. Please connect first."; return }
    Write-Host "Extending partition using diskpart..."
    $vmName = Read-Host "Please enter the name of the VM to configure"
    try {
        $vm = Get-VM -Name $vmName -ErrorAction Stop
        $username = "$vmName\ctxadmin"
        $password = "Banorte2020."
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        $guestCredential = New-Object System.Management.Automation.PSCredential($username, $securePassword)
        if ($vm.PowerState -ne "PoweredOn") {
            Write-Host "Powering on VM '$vmName'..."
            Start-VM -VM $vm -Confirm:$false
            Start-Sleep -Seconds 60
        }
        $diskpartCommands = @"
select disk 0
select partition 3
delete partition override
select partition 2
extend
exit
"@
        $guestScript = @"
`$diskpartScriptPath = 'C:\Windows\Temp\diskpart_script.txt'
Set-Content -Path `$diskpartScriptPath -Value @'
$diskpartCommands
'@
diskpart /s `$diskpartScriptPath
if (`$LASTEXITCODE -eq 0) { Write-Output 'Diskpart commands executed successfully. Partition 3 deleted and partition 2 extended.' }
else { Write-Output 'Error executing diskpart commands.'; exit 1 }
Remove-Item -Path `$diskpartScriptPath -Force
"@
        Write-Host "Running diskpart to extend partition on VM '$vmName'..."
        $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
        Write-Host "Script output: $($result.ScriptOutput)"
        Write-Host "Partition extension completed for VM '$vmName'."
    } catch { Write-Host "Error extending partition: $_" }
}

# --- Main Script Logic ---
do {
    $mainChoice = Show-MainMenu
    switch ($mainChoice) {
        "1" { # Snapshot Management
            $vCenterServer = Select-vCenter
            $credential = Get-Credential
            try {
                Connect-VIServer -Server $vCenterServer -Credential $credential -ErrorAction Stop
                Log-Message "Connected to vCenter: $vCenterServer"
            } catch {
                Log-Message "Failed to connect to vCenter: $vCenterServer. Error: $_"
                continue
            }
            do {
                $snapChoice = Show-SnapshotMenu
                switch ($snapChoice) {
                    "1" {
                        while ($true) {
                            Write-Host "Select an option to choose VMs:"
                            Write-Host "1. Provide a CSV list with VM names"
                            Write-Host "2. Select VMs from a popup window list"
                            $option = Read-Host "Enter your choice (1 or 2)"
                            if ($option -in 1..2) { break }
                            Write-Host "Invalid choice. Please enter 1 or 2."
                        }
                        if ($option -eq 1) {
                            $csvPath = Get-CSVFilePath
                            try { $vmNames = Import-Csv -Path $csvPath | Select-Object -ExpandProperty VMName -ErrorAction Stop }
                            catch { Log-Message "Failed to read CSV file: $_"; continue }
                        } else { $vmNames = Select-VMsFromList }
                        if ($vmNames.Count -gt 0) {
                            $snapshotName = Read-Host "Enter the snapshot name"
                            $snapshotDescription = Read-Host "Enter the snapshot description"
                            Take-VMSnapshot -VMNames $vmNames -SnapshotName $snapshotName -SnapshotDescription $snapshotDescription
                        } else { Log-Message "No VMs selected." }
                    }
                    "2" {
                        $option = Read-Host "Select VMs from CSV (1) or list (2)?"
                        $vmNames = if ($option -eq "1") { Import-Csv -Path (Get-CSVFilePath) | Select-Object -ExpandProperty VMName } else { Select-VMsFromList }
                        if ($vmNames.Count -gt 0) { Remove-LastSnapshot -VMNames $vmNames }
                    }
                    "3" {
                        $option = Read-Host "Select VMs from CSV (1) or list (2)?"
                        $vmNames = if ($option -eq "1") { Import-Csv -Path (Get-CSVFilePath) | Select-Object -ExpandProperty VMName } else { Select-VMsFromList }
                        if ($vmNames.Count -gt 0) { Remove-AllSnapshots -VMNames $vmNames }
                    }
                    "4" {
                        $option = Read-Host "Select VMs from CSV (1) or list (2)?"
                        $vmNames = if ($option -eq "1") { Import-Csv -Path (Get-CSVFilePath) | Select-Object -ExpandProperty VMName } else { Select-VMsFromList }
                        if ($vmNames.Count -gt 0) { Remove-SelectedSnapshots -VMNames $vmNames }
                    }
                    "5" {
                        $option = Read-Host "Select VMs from CSV (1) or list (2)?"
                        $vmNames = if ($option -eq "1") { Import-Csv -Path (Get-CSVFilePath) | Select-Object -ExpandProperty VMName } else { Select-VMsFromList }
                        if ($vmNames.Count -gt 0) { Restart-GuestOS -VMNames $vmNames }
                    }
                    "6" { break }
                    default { Write-Host "Invalid option." }
                }
                $anotherAction = Read-Host "Do you want to perform another action? (yes/no)"
            } while ($anotherAction -eq "yes")
            Disconnect-VIServer -Confirm:$false
            Log-Message "Disconnected from vCenter: $vCenterServer"
        }
        "2" { # Server Management
            Add-Type -AssemblyName System.Windows.Forms
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
            if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
                Write-Host "No CSV file selected. Skipping Server Management..."
                continue
            }
            $csvPath = $openFileDialog.FileName
            $servers = Import-Csv -Path $csvPath
            do {
                $serverChoice = Show-ServerMenu
                switch ($serverChoice) {
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
                    "14" { break }
                    default { Write-Host "Invalid option." }
                }
                if ($serverChoice -ne "14") { $another = Read-Host "Do you want to perform another action? (Y/N)" }
            } while (($another -eq "Y" -or $another -eq "y") -and $serverChoice -ne "14")
        }
        "3" { # VM Management
            do {
                $vmChoice = Show-VMMenu
                switch ($vmChoice) {
                    "1" {
                        Write-Host "Connecting to vCenter..."
                        $vCenterServer = Read-Host "Please enter the vCenter server address"
                        $vCenterCredential = Get-Credential -Message "Enter your vCenter credentials"
                        try {
                            Connect-VIServer -Server $vCenterServer -Credential $vCenterCredential -ErrorAction Stop
                            Write-Host "Successfully connected to vCenter: $vCenterServer"
                        } catch { Write-Host "Error connecting to vCenter: $_" }
                    }
                    "2" { Clone-VM }
                    "3" { Disconnect-vNIC }
                    "4" { Connect-vNIC }
                    "5" { Rename-Computer }
                    "6" { Set-CtxAdmin2Password }
                    "7" { Start-VMInstance }
                    "8" { Extend-Partition }
                    "9" { break }
                    default { Write-Host "Invalid option. Please select a number between 1 and 9." }
                }
                if ($vmChoice -ne "9") { $continue = (Read-Host "Do you want to perform another action? (yes/no)") -eq "yes" }
            } while ($continue)
            if ($global:DefaultVIServer) {
                Disconnect-VIServer -Server $global:DefaultVIServer -Confirm:$false
                Write-Host "Disconnected from vCenter."
            }
        }
        "4" { Write-Host "Exiting script..."; break }
        default { Write-Host "Invalid option." }
    }
} while ($true)