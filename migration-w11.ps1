# Function to display the menu
function Show-Menu {
    Clear-Host
    Write-Host "=================================="
    Write-Host "PowerCLI VM Management Menu"
    Write-Host "=================================="
    Write-Host "1. Connect to vCenter"
    Write-Host "2. Clone VM"
    Write-Host "3. Disconnect vNIC"
    Write-Host "4. Connect vNIC"
    Write-Host "5. Rename Computer"
    Write-Host "6. Set Password for CtxAdmin2"
    Write-Host "7. Start VM"
    Write-Host "8. Extend Partition"
    Write-Host "9. Quit"
    Write-Host "=================================="
    
    # Show connection status
    if ($global:DefaultVIServer) {
        Write-Host "Connected to: $($global:DefaultVIServer.Name)" -ForegroundColor Green
    } else {
        Write-Host "Not connected to vCenter" -ForegroundColor Yellow
    }
    Write-Host "=================================="
    
    $choice = Read-Host "Please select an option (1-9)"
    return $choice
}

# Function to pause and wait for user input before returning to menu
function Wait-ForUser {
    param([string]$Message = "Press any key to return to menu...")
    Write-Host ""
    Write-Host $Message -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Function to validate vCenter connection
function Test-VCenterConnection {
    if (-not $global:DefaultVIServer) {
        Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)." -ForegroundColor Red
        Wait-ForUser
        return $false
    }
    return $true
}

# Main script loop
$continue = $true
while ($continue) {
    $choice = Show-Menu

    switch ($choice) {
        "1" {
            # 1. Connect to vCenter
            Write-Host "Connecting to vCenter..." -ForegroundColor Yellow
            $vCenterServer = Read-Host "Please enter the vCenter server address"
            $vCenterCredential = Get-Credential -Message "Enter your vCenter credentials"
            try {
                Connect-VIServer -Server $vCenterServer -Credential $vCenterCredential -ErrorAction Stop
                Write-Host "Successfully connected to vCenter: $vCenterServer" -ForegroundColor Green
            }
            catch {
                Write-Host "Error connecting to vCenter: $_" -ForegroundColor Red
            }
            Wait-ForUser
        }

        "2" {
            # 2. Clone VM
            if (-not (Test-VCenterConnection)) { continue }

            Write-Host "Cloning VM..." -ForegroundColor Yellow
            $sourceVMName = Read-Host "Please enter the name of the source VM to clone"
            $newVMName = Read-Host "Please enter the name for the new VM"

            try {
                $sourceVM = Get-VM -Name $sourceVMName -ErrorAction Stop
                $newVM = New-VM -VM $sourceVM -Name $newVMName -VMHost $sourceVM.VMHost -Datastore "VMsRecursosTercerosHA" -ErrorAction Stop
                Write-Host "VM '$newVMName' has been cloned successfully from '$sourceVMName'." -ForegroundColor Green
            }
            catch {
                Write-Host "Error cloning VM: $_" -ForegroundColor Red
            }
            Wait-ForUser
        }

        "3" {
            # 3. Disconnect vNIC
            if (-not (Test-VCenterConnection)) { continue }

            Write-Host "Disconnecting vNIC..." -ForegroundColor Yellow
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop
                $nic = Get-NetworkAdapter -VM $vm -ErrorAction Stop
                Get-NetworkAdapter -VM $vm | Set-NetworkAdapter -NetworkName "VDI_Terceros_2067" -StartConnected:$false -Confirm:$false -ErrorAction Stop
                Write-Host "VM '$vmName' has been configured with VLAN 'VDI_Terceros_2067' and disconnected network adapter." -ForegroundColor Green
            }
            catch {
                Write-Host "Error disconnecting vNIC: $_" -ForegroundColor Red
            }
            Wait-ForUser
        }

        "4" {
            # 4. Connect vNIC
            if (-not (Test-VCenterConnection)) { continue }

            Write-Host "Connecting vNIC..." -ForegroundColor Yellow
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop
                Get-NetworkAdapter -VM $vm | Set-NetworkAdapter -Connected:$true -Confirm:$false -ErrorAction Stop
                $updatedAdapter = Get-NetworkAdapter -VM $vm
                if ($updatedAdapter.ConnectionState.Connected) {
                    Write-Host "Success: vNIC for VM '$vmName' is now connected to VLAN '$($updatedAdapter.NetworkName)'." -ForegroundColor Green
                }
                else {
                    Write-Host "Error: vNIC for VM '$vmName' is still disconnected." -ForegroundColor Red
                }
            }
            catch {
                Write-Host "Error connecting vNIC: $_" -ForegroundColor Red
            }
            Wait-ForUser
        }

        "5" {
            # 5. Rename Computer
            if (-not (Test-VCenterConnection)) { continue }

            Write-Host "Renaming computer..." -ForegroundColor Yellow
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop

                # Construct the local admin username in the format vmname\ctxadmin
                $username = "$vmName\ctxadmin"
                $password = "Banorte2020."
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                $guestCredential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

                $newComputerName = Read-Host "Please enter the new computer name for VM '$vmName'"
                $workgroupName = Read-Host "Please enter the workgroup name for the VM (e.g., WORKGROUP)"

                if ($vm.PowerState -ne "PoweredOn") {
                    Write-Host "Powering on VM '$vmName'..." -ForegroundColor Yellow
                    Start-VM -VM $vm -Confirm:$false
                    Write-Host "Waiting for VM to boot up..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 60
                }

                $guestScript = @"
                Remove-Computer -Force -PassThru
                Add-Computer -WorkgroupName "$workgroupName" -Force
                Rename-Computer -NewName "$newComputerName" -Force
                Write-Output "Computer name set to $newComputerName and joined workgroup $workgroupName. Rebooting now..."
                Restart-Computer -Force
"@

                Write-Host "Applying computer name and workgroup changes to VM '$vmName'..." -ForegroundColor Yellow
                $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
                Write-Host "Script output: $($result.ScriptOutput)" -ForegroundColor Cyan
                Write-Host "VM '$vmName' has been configured with the new computer name '$newComputerName' and workgroup '$workgroupName'. The VM is rebooting to apply changes." -ForegroundColor Green
                Write-Host "Note: The computer object for '$vmName' may still exist in Active Directory. Please manually remove it from the domain to avoid stale entries." -ForegroundColor Yellow
            }
            catch {
                Write-Host "Error renaming computer: $_" -ForegroundColor Red
            }
            Wait-ForUser
        }

        "6" {
            # 6. Set Password for CtxAdmin2
            if (-not (Test-VCenterConnection)) { continue }

            Write-Host "Setting password for ctxadmin2..." -ForegroundColor Yellow
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop

                # Construct the local admin username in the format vmname\ctxadmin
                $username = "$vmName\ctxadmin"
                $password = "Banorte2020."
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                $guestCredential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

                if ($vm.PowerState -ne "PoweredOn") {
                    Write-Host "Powering on VM '$vmName'..." -ForegroundColor Yellow
                    Start-VM -VM $vm -Confirm:$false
                    Write-Host "Waiting for VM to boot up..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 60
                }

                $newPassword = "VNTB4n0rt3B4n0rt3"
                $guestScript = @"
                try {
                    `$user = [ADSI]('WinNT://./ctxadmin2,user')
                    if (`$user -eq `$null) {
                        Write-Output 'Error: User ctxadmin2 not found.'
                        exit 1
                    }
                    `$user.SetPassword('$newPassword')
                    `$user.SetInfo()
                    Write-Output 'Password for ctxadmin2 changed successfully.'
                }
                catch {
                    Write-Output 'Error changing password for ctxadmin2: `$_'
                    exit 1
                }
"@

                Write-Host "Changing password for ctxadmin2 on VM '$vmName'..." -ForegroundColor Yellow
                $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
                Write-Host "Script output: $($result.ScriptOutput)" -ForegroundColor Cyan
                Write-Host "Password change completed for user ctxadmin2 on VM '$vmName'." -ForegroundColor Green
            }
            catch {
                Write-Host "Error setting password for ctxadmin2: $_" -ForegroundColor Red
            }
            Wait-ForUser
        }

        "7" {
            # 7. Start VM
            if (-not (Test-VCenterConnection)) { continue }

            Write-Host "Starting VM..." -ForegroundColor Yellow
            $vmName = Read-Host "Please enter the name of the VM to start"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop
                if ($vm.PowerState -eq "PoweredOn") {
                    Write-Host "VM '$vmName' is already powered on." -ForegroundColor Yellow
                } else {
                    Start-VM -VM $vm -Confirm:$false -ErrorAction Stop
                    Write-Host "VM '$vmName' has been started." -ForegroundColor Green
                }
            }
            catch {
                Write-Host "Error starting VM: $_" -ForegroundColor Red
            }
            Wait-ForUser
        }

        "8" {
            # 8. Extend Partition
            if (-not (Test-VCenterConnection)) { continue }

            Write-Host "Extending partition using diskpart..." -ForegroundColor Yellow
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop

                # Construct the local admin username in the format vmname\ctxadmin
                $username = "$vmName\ctxadmin"
                $password = "Banorte2020."
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                $guestCredential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

                if ($vm.PowerState -ne "PoweredOn") {
                    Write-Host "Powering on VM '$vmName'..." -ForegroundColor Yellow
                    Start-VM -VM $vm -Confirm:$false
                    Write-Host "Waiting for VM to boot up..." -ForegroundColor Yellow
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

                Write-Host "Running diskpart to extend partition on VM '$vmName'..." -ForegroundColor Yellow
                $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
                Write-Host "Script output: $($result.ScriptOutput)" -ForegroundColor Cyan
                Write-Host "Partition extension completed for VM '$vmName'." -ForegroundColor Green
            }
            catch {
                Write-Host "Error extending partition: $_" -ForegroundColor Red
            }
            Wait-ForUser
        }

        "9" {
            # 9. Quit
            Write-Host "Exiting script..." -ForegroundColor Yellow
            $continue = $false
            continue
        }

        default {
            Write-Host "Invalid option. Please select a number between 1 and 9." -ForegroundColor Red
            Wait-ForUser
        }
    }
}

# Disconnect from vCenter if connected
if ($global:DefaultVIServer) {
    Disconnect-VIServer -Server $global:DefaultVIServer -Confirm:$false
    Write-Host "Disconnected from vCenter." -ForegroundColor Green
}