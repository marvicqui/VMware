# Function to display the menu
function Show-Menu {
    Write-Host "=================================="
    Write-Host "PowerCLI VM Management Menu"
    Write-Host "=================================="
    Write-Host "1. Connect to vCenter"
    Write-Host "2. Clone VM"
    Write-Host "3. Disconnect vNIC"
    Write-Host "4. Connect vNIC"
    Write-Host "5. Rename Computer"
    Write-Host "6. Configure Admin Users (ctxadmin2 & ugsop_vdiha)"
    Write-Host "7. Start VM"
    Write-Host "8. Extend Partition"
    Write-Host "9. Snipping Tools Installation"
    Write-Host "10. Quit"
    Write-Host "=================================="
    $choice = Read-Host "Please select an option (1-10)"
    return $choice
}

# Main script loop
$continue = $true
while ($continue) {
    $choice = Show-Menu

    switch ($choice) {
        "1" {
            # 1. Connect to vCenter
            Write-Host "Connecting to vCenter..."
            #$vCenterServer = Read-Host "Please enter the vCenter server address"
            $vCenterServer = "vmwvcsttapha01.edificios.gfbanorte"
            #$vCenterCredential = Get-Credential -Message "Enter your vCenter credentials"
            try {
                #Connect-VIServer -Server $vCenterServer -Credential $vCenterCredential -ErrorAction Stop
                $vcenterConnection = Connect-VIServer -Server $vCenterServer -user hviqm800@vsphere.local -pass Passw0rd.1 -ErrorAction Stop
                Write-Host "Successfully connected to vCenter: $vCenterServer"
            }
            catch {
                Write-Host "Error connecting to vCenter: $_"
            }
        }

        "2" {
            # 2. Clone VM
            if (-not $global:DefaultVIServer) {
                Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)."
                continue
            }

            Write-Host "Cloning VM..."
            $sourceVMName = Read-Host "Please enter the name of the source VM to clone"
            $newVMName = Read-Host "Please enter the name for the new VM"

            try {
                $sourceVM = Get-VM -Name $sourceVMName -ErrorAction Stop
                $newVM = New-VM -VM $sourceVM -Name $newVMName -VMHost $sourceVM.VMHost -Datastore "VMsRecursosTercerosHA" -ErrorAction Stop
                Write-Host "VM '$newVMName' has been cloned successfully from '$sourceVMName'."
            }
            catch {
                Write-Host "Error cloning VM: $_"
            }
        }

        "3" {
            # 3. Disconnect vNIC
            if (-not $global:DefaultVIServer) {
                Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)."
                continue
            }

            Write-Host "Disconnecting vNIC..."
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop
                $nic = Get-NetworkAdapter -VM $vm -ErrorAction Stop
                Get-NetworkAdapter -VM $vm | Set-NetworkAdapter -NetworkName "VDI_Terceros_2067" -StartConnected:$false -Confirm:$false -ErrorAction Stop
                Write-Host "VM '$vmName' has been configured with VLAN 'VDI_Terceros_2067' and disconnected network adapter."
            }
            catch {
                Write-Host "Error disconnecting vNIC: $_"
            }
        }

        "4" {
            # 4. Connect vNIC
            if (-not $global:DefaultVIServer) {
                Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)."
                continue
            }

            Write-Host "Connecting vNIC..."
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop
                Get-NetworkAdapter -VM $vm | Set-NetworkAdapter -Connected:$true -Confirm:$false -ErrorAction Stop
                $updatedAdapter = Get-NetworkAdapter -VM $vm
                if ($updatedAdapter.ConnectionState.Connected) {
                    Write-Host "Success: vNIC for VM '$vmName' is now connected to VLAN '$($updatedAdapter.NetworkName)'."
                }
                else {
                    Write-Host "Error: vNIC for VM '$vmName' is still disconnected."
                }
            }
            catch {
                Write-Host "Error connecting vNIC: $_"
            }
        }

        "5" {
            # 5. Rename Computer
            if (-not $global:DefaultVIServer) {
                Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)."
                continue
            }

            Write-Host "Renaming computer..."
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop

                # Construct the local admin username in the format vmname\ctxadmin
                $username = ".\ctxadmin"
                $password = "Banorte2020."
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                $guestCredential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

                $newComputerName = $vmName
                
                if ($vm.PowerState -ne "PoweredOn") {
                    Write-Host "Powering on VM '$vmName'..."
                    Start-VM -VM $vm -Confirm:$false
                    Start-Sleep -Seconds 60
                }

                $guestScript = @"
                Remove-Computer -Force -PassThru
                Rename-Computer -NewName "$newComputerName" -Force
                Write-Output "Computer name set to $newComputerName and joined workgroup $workgroupName. Rebooting now..."
                Restart-Computer -Force
"@

                Write-Host "Applying computer name and workgroup changes to VM '$vmName'..."
                $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
                Write-Host "Script output: $($result.ScriptOutput)"
                Write-Host "VM '$vmName' has been configured with the new computer name '$newComputerName'. The VM is rebooting to apply changes."
                Write-Host "Note: The computer object for '$vmName' may still exist in Active Directory. Please manually remove it from the domain to avoid stale entries."
            }
            catch {
                Write-Host "Error renaming computer: $_"
            }
        }

        "6" {
            # 6. Configure Admin Users (ctxadmin2 & ugsop_vdiha)
            if (-not $global:DefaultVIServer) {
                Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)."
                continue
            }

            Write-Host "Configuring admin users..."
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop

                # Construct the local admin username for authentication
                $username = ".\ctxadmin"
                $password = "Banorte2020."
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                $guestCredential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

                if ($vm.PowerState -ne "PoweredOn") {
                    Write-Host "Powering on VM '$vmName'..."
                    Start-VM -VM $vm -Confirm:$false
                    Start-Sleep -Seconds 60
                }

                # Define passwords and usernames
                $ctxadmin2Password = "VNTB4n0rt3B4n0rt3"
                $ugsopUsername = "ugsop_vdiha"
                $ugsopPassword = "VNTB4n0rt3B4n0rt3"
                $adminGroupName = "Administradores"

                # Combined guest script
                $guestScript = @"
                # Task 1: Set password for ctxadmin2
                try {
                    `$user2 = [ADSI]('WinNT://./ctxadmin2,user')
                    if (`$user2 -eq `$null) {
                        Write-Output 'Info: User ctxadmin2 not found. Skipping password change.'
                    }
                    else {
                        `$user2.SetPassword('$ctxadmin2Password')
                        `$user2.SetInfo()
                        Write-Output 'Password for ctxadmin2 changed successfully.'
                    }
                }
                catch {
                    Write-Output 'Error changing password for ctxadmin2: `$_'
                }

                # Task 2: Create and configure ugsop_vdiha
                try {
                    # Check if the user already exists
                    if (-not (Get-LocalUser -Name '$ugsopUsername' -ErrorAction SilentlyContinue)) {
                        New-LocalUser -Name '$ugsopUsername' -Password (ConvertTo-SecureString '$ugsopPassword' -AsPlainText -Force) -FullName 'VDI HA User' -Description 'Local admin account for VDI HA'
                        Write-Output 'User $ugsopUsername created successfully.'
                        Start-Sleep -Seconds 2
                    } else {
                        Write-Output 'User $ugsopUsername already exists.'
                    }

                    # Configure user
                    Set-LocalUser -Name '$ugsopUsername' -PasswordNeverExpires `$true
                    Write-Output 'Password for $ugsopUsername set to never expire.'

                    if (-not (Get-LocalGroupMember -Group '$adminGroupName' -Member '$ugsopUsername' -ErrorAction SilentlyContinue)) {
                        Add-LocalGroupMember -Group '$adminGroupName' -Member '$ugsopUsername'
                        Write-Output 'User $ugsopUsername added to the $adminGroupName group.'
                    } else {
                        Write-Output 'User $ugsopUsername is already a member of the $adminGroupName group.'
                    }
                }
                catch {
                    Write-Output 'Error processing user ugsop_vdiha: `$_'
                }
"@

                Write-Host "Running user configuration script on VM '$vmName'..."
                $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
                Write-Host "Script output: $($result.ScriptOutput)"
                Write-Host "User configuration completed on VM '$vmName'."
            }
            catch {
                Write-Host "Error configuring users: $_"
            }
        }

        "7" {
            # 7. Start VM
            if (-not $global:DefaultVIServer) {
                Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)."
                continue
            }

            Write-Host "Starting VM..."
            $vmName = Read-Host "Please enter the name of the VM to start"

            try {
                Start-VM -VM $vmName -Confirm:$false -ErrorAction Stop
                Write-Host "VM '$vmName' has been started."
            }
            catch {
                Write-Host "Error starting VM: $_"
            }
        }

        "8" {
            # 8. Extend Partition
            if (-not $global:DefaultVIServer) {
                Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)."
                continue
            }

            Write-Host "Extending partition using diskpart..."
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop

                # Construct the local admin username in the format vmname\ctxadmin
                $username = "$vmName\ctxadmin"
                $password = "Banorte2020."
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                $guestCredential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

                if ($vm.PowerState -ne "PoweredOn") {
                    Write-Host "Powering on VM '$vmName'..."
                    Start-VM -VM $vm -Confirm:$false
                    Start-Sleep -Seconds 60
                }

                # Create a diskpart script file in the guest OS
                $diskpartCommands = @"
select disk 0
select partition 3
delete partition override
select partition 2
extend
exit
"@

                $guestScript = @"
# Write the diskpart commands to a temporary file
try {
    `$diskpartScriptPath = 'C:\Windows\Temp\diskpart_script.txt'
    Set-Content -Path `$diskpartScriptPath -Value @'
$diskpartCommands
'@

    # Run diskpart with the script file
    diskpart /s `$diskpartScriptPath

    # Check if diskpart executed successfully
    if (`$LASTEXITCODE -eq 0) {
        Write-Output 'Diskpart commands executed successfully. Partition 3 deleted and partition 2 extended.'
    }
    else {
        Write-Output 'Error executing diskpart commands.'
        exit 1
    }

    # Clean up the temporary script file
    Remove-Item -Path `$diskpartScriptPath -Force
}
catch {
    Write-Output 'Error executing diskpart: `$_'
    exit 1
}
"@

                Write-Host "Running diskpart to extend partition on VM '$vmName'..."
                $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
                Write-Host "Script output: $($result.ScriptOutput)"
                Write-Host "Partition extension completed for VM '$vmName'."
            }
            catch {
                Write-Host "Error extending partition: $_"
            }
        }

        "9" {
            # 9. Snipping Tools Installation
            if (-not $global:DefaultVIServer) {
                Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)."
                continue
            }

            Write-Host "Installing Snipping Tool and dependencies..."
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop

                # Construct the local admin username for authentication
                $username = ".\ctxadmin"
                $password = "Banorte2020."
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                $guestCredential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

                if ($vm.PowerState -ne "PoweredOn") {
                    Write-Host "Powering on VM '$vmName'..."
                    Start-VM -VM $vm -Confirm:$false
                    Start-Sleep -Seconds 60
                }

                # This script block will run inside the guest VM
                $guestScript = @"
                `$ErrorActionPreference = 'Stop'
                `$basePath = 'C:\Temporal\recorte'

                Write-Output 'Starting installation of Snipping Tool and dependencies...'
                Write-Output "Searching for files in: `$basePath"

                # Check if base directory exists
                if (-not (Test-Path -Path `$basePath -PathType Container)) {
                    Write-Output "FATAL: Directory `$basePath does not exist on the VM. Aborting."
                    exit 1
                }

                # List of required packages in order of installation
                `$packages = @(
                    'Microsoft.WindowsAppRuntime.1.5_5001.275.500.0_Universal_X64.msix',
                    'Microsoft.UI.Xaml.2.8_8.2310.30001.0_Universal_X64.appx',
                    'Microsoft.VCLibs.140.00_14.0.33519.0_Universal_X64.appx',
                    'Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_Universal_X64.appx',
                    'Microsoft.ScreenSketch_2022.2407.3.0_Desktop_Arm64.X64.msixbundle'
                )

                # Install each package
                foreach (`$pkgFile in `$packages) {
                    `$fullPath = Join-Path -Path `$basePath -ChildPath `$pkgFile
                    Write-Output "---"
                    Write-Output "Attempting to install: `$pkgFile"

                    if (-not (Test-Path `$fullPath -PathType Leaf)) {
                        Write-Output "ERROR: File not found - `$fullPath"
                        continue # Skip to the next file
                    }

                    try {
                        Add-AppxPackage -Path `$fullPath
                        Write-Output "SUCCESS: Installed `$pkgFile"
                    } catch {
                        Write-Output "ERROR: Failed to install `$pkgFile. Details: `$_.Exception.Message"
                    }
                }

                # Provision the main package for all users
                Write-Output '---'
                Write-Output "Attempting to provision Screen Sketch package..."
                `$screenSketchPath = Join-Path -Path `$basePath -ChildPath 'Microsoft.ScreenSketch_2022.2407.3.0_Desktop_Arm64.X64.msixbundle'
                if (Test-Path `$screenSketchPath -PathType Leaf) {
                    try {
                        Add-AppxProvisionedPackage -Online -PackagePath `$screenSketchPath -SkipLicense
                        Write-Output 'SUCCESS: Screen Sketch provisioned for all new users.'
                    } catch {
                        Write-Output "ERROR: Failed to provision Screen Sketch. Details: `$_.Exception.Message"
                    }
                } else {
                    Write-Output "ERROR: Screen Sketch package not found for provisioning."
                }

                # Disable the scheduled task
                Write-Output '---'
                Write-Output 'Attempting to disable PcaWallpaperAppDetect scheduled task...'
                try {
                    Disable-ScheduledTask -TaskPath '\Microsoft\Windows\Application Experience\' -TaskName 'PcaWallpaperAppDetect'
                    Write-Output 'SUCCESS: Scheduled task "PcaWallpaperAppDetect" disabled.'
                } catch {
                    if (`$_.Exception.Message -like '*The system cannot find the file specified*') {
                        Write-Output 'INFO: Scheduled task "PcaWallpaperAppDetect" not found. No action needed.'
                    } else {
                        Write-Output "ERROR: Failed to disable scheduled task. Details: `$_.Exception.Message"
                    }
                }

                Write-Output '---'
                Write-Output 'Installation script finished.'
"@

                Write-Host "Running Snipping Tool installation script on VM '$vmName'..."
                $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
                Write-Host "Guest script output:"
                # Format the output from the guest
                $result.ScriptOutput.Split([Environment]::NewLine) | ForEach-Object { Write-Host "  $_" }
                Write-Host "Snipping Tool installation process completed on VM '$vmName'."
            }
            catch {
                Write-Host "An error occurred during the Snipping Tool installation process: $_"
            }
        }

        "10" {
            # 10. Quit
            Write-Host "Exiting script..."
            $continue = $false
            continue
        }

        default {
            Write-Host "Invalid option. Please select a number between 1 and 10."
        }
    }
}

# Disconnect from vCenter if connected
if ($global:DefaultVIServer) {
    Disconnect-VIServer -Server $global:DefaultVIServer -Confirm:$false
    Write-Host "Disconnected from vCenter."
}