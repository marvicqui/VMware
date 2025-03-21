# Function to display the menu
function Show-Menu {
    Write-Host "=================================="
    Write-Host "PowerCLI VM Management Menu"
    Write-Host "=================================="
    Write-Host "1. Connect to vCenter"
    Write-Host "2. Clone VM"
    Write-Host "3. Set VLAND2067 and disconnect vNIC"
    Write-Host "4. Connect vNIC"
    Write-Host "5. Rename Computer"
    Write-Host "6. Set Password for CtxAdmin2"
    Write-Host "7. Start VM"
    Write-Host "8. Quit"
    Write-Host "=================================="
    $choice = Read-Host "Please select an option (1-8)"
    return $choice
}

# Function to ask if the user wants to perform another action
function Ask-Continue {
    $continue = Read-Host "Do you want to perform another action? (yes/no)"
    return $continue -eq "yes"
}

# Main script loop
$continue = $true
while ($continue) {
    $choice = Show-Menu

    switch ($choice) {
        "1" {
            # 1. Connect to vCenter
            Write-Host "Connecting to vCenter..."
            $vCenterServer = Read-Host "Please enter the vCenter server address"
            $vCenterCredential = Get-Credential -Message "Enter your vCenter credentials"
            try {
                Connect-VIServer -Server $vCenterServer -Credential $vCenterCredential -ErrorAction Stop
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

            Write-Host "Disconnecting vNIC and setting VLAN VDI_Terceros_2067..."
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
            $vmName = Read-Host "Please enter the name of the VM to configure, it will be the new computer name"

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
                Write-Output "Computer name set to $newComputerName. Rebooting now..."
                Restart-Computer -Force
"@

                Write-Host "Applying computer name and workgroup changes to VM '$vmName'..."
                $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
                Write-Host "Script output: $($result.ScriptOutput)"
                Write-Host "VM '$vmName' has been configured with the new computer name '$newComputerName' '. The VM is rebooting to apply changes."
                
            }
            catch {
                Write-Host "Error renaming computer: $_"
            }
        }

        "6" {
            # 6. Set Password for CtxAdmin2
            if (-not $global:DefaultVIServer) {
                Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)."
                continue
            }

            Write-Host "Setting password for ctxadmin2..."
            $vmName = Read-Host "Please enter the name of the VM to configure"

            try {
                $vm = Get-VM -Name $vmName -ErrorAction Stop

                # Construct the local admin username in the format vmname\ctxadmin
                $username = ".\ctxadmin"
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

                Write-Host "Changing password for ctxadmin2 on VM '$vmName'..."
                $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
                Write-Host "Script output: $($result.ScriptOutput)"
                Write-Host "Password change completed for user ctxadmin2 on VM '$vmName'."
            }
            catch {
                Write-Host "Error setting password for ctxadmin2: $_"
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
            # 8. Quit
            Write-Host "Exiting script..."
            $continue = $false
            continue
        }

        default {
            Write-Host "Invalid option. Please select a number between 1 and 8."
        }
    }

    # Ask if the user wants to perform another action
    if ($choice -ne "8") {
        $continue = Ask-Continue
    }
}

# Disconnect from vCenter if connected
if ($global:DefaultVIServer) {
    Disconnect-VIServer -Server $global:DefaultVIServer -Confirm:$false
    Write-Host "Disconnected from vCenter."
}