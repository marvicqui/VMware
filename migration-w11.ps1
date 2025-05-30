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
    Write-Host "6. Set Password for CtxAdmin2"
    Write-Host "7. Start VM"
    Write-Host "8. Extend Partition"
    Write-Host "9. Create Local User ugsop_vdiha"
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
            # 9. Create Local User ugsop_vdiha
            if (-not $global:DefaultVIServer) {
                Write-Host "Error: Not connected to a vCenter Server. Please connect first (option 1)."
                continue
            }

            Write-Host "Creating local user ugsop_vdiha..."
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

                $newUsername = "ugsop_vdiha"
                $newPassword = "VNTB4n0rt3B4n0rt3"
                $groupName = "Administradores"

                $guestScript = @"
                try {
                    # Check if the user already exists
                    if (-not (Get-LocalUser -Name '$newUsername' -ErrorAction SilentlyContinue)) {
                        # Create a new local user
                        New-LocalUser -Name '$newUsername' -Password (ConvertTo-SecureString '$newPassword' -AsPlainText -Force) -FullName 'VDI HA User' -Description 'Local admin account for VDI HA'
                        Write-Output 'User $newUsername created successfully.'
                        # Add a slight delay to ensure the user is fully committed
                        Start-Sleep -Seconds 2
                    } else {
                        Write-Output 'User $newUsername already exists.'
                    }

                    # Set the password to never expire (always apply, whether user is new or existing)
                    Set-LocalUser -Name '$newUsername' -PasswordNeverExpires `$true
                    Write-Output 'Password for $newUsername set to never expire.'

                    # Add the user to the Administrators group
                    if (-not (Get-LocalGroupMember -Group '$groupName' -Member '$newUsername' -ErrorAction SilentlyContinue)) {
                        Add-LocalGroupMember -Group '$groupName' -Member '$newUsername'
                        Write-Output 'User $newUsername added to the $groupName group.'
                    } else {
                        Write-Output 'User $newUsername is already a member of the $groupName group.'
                    }
                }
                catch {
                    Write-Output 'Error creating user, setting password to never expire, or adding to group: `$_'
                    exit 1
                }
"@

                Write-Host "Creating user $newUsername on VM '$vmName'..."
                $result = Invoke-VMScript -VM $vm -ScriptText $guestScript -GuestCredential $guestCredential -ScriptType PowerShell -ErrorAction Stop
                Write-Host "Script output: $($result.ScriptOutput)"
                Write-Host "User creation completed for $newUsername on VM '$vmName'."
            }
            catch {
                Write-Host "Error creating local user: $_"
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