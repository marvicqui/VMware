<#
.SYNOPSIS
    This script connects to a specified vCenter server and performs snapshot operations on selected virtual machines (VMs).

.DESCRIPTION
    The script allows the user to connect to one of the predefined vCenter servers, select VMs either from a CSV file or a popup window, and perform one of the following actions:
    1. Take snapshots of the selected VMs.
    2. Delete the last snapshot of the selected VMs.
    3. Delete all snapshots of the selected VMs.

.PARAMETER vCenters
    An array of vCenter server addresses to choose from.

.PARAMETER Select-vCenter
    Prompts the user to select a vCenter server from the predefined list.

.PARAMETER Get-Credential
    Prompts the user to enter credentials for connecting to the selected vCenter server.

.PARAMETER Connect-VIServer
    Connects to the selected vCenter server using the provided credentials.

.PARAMETER Take-VMSnapshot
    Takes a snapshot of the specified VMs with the provided snapshot name and description.

.PARAMETER Remove-LastSnapshot
    Deletes the last snapshot of the specified VMs.

.PARAMETER Remove-AllSnapshots
    Deletes all snapshots of the specified VMs.

.PARAMETER Get-CSVFilePath
    Opens a file explorer dialog to select a CSV file containing VM names.

.PARAMETER Select-VMsFromList
    Opens a popup window to select VMs from a list of available VMs.

.PARAMETER Main script logic
    Prompts the user to select an action to perform (take snapshots, delete the last snapshot, or delete all snapshots) and a method to choose VMs (CSV file or popup window).

.NOTES
    Requires VMware PowerCLI module to be installed and imported.
    The script will exit if an invalid selection is made or if no VMs are selected.

.EXAMPLE
    # Run the script and follow the prompts to connect to a vCenter server, select VMs, and perform snapshot operations.
    .\YourScriptName.ps1
#>
# Set-PowerCLIConfiguration -Scope User -ParticipateInCeip $false -Confirm:$false
# Import VMware PowerCLI module
Import-Module VMware.PowerCLI

# List of vCenters
$vCenters = @(
    "vwvsinap01.edificios.gfbanorte",
    "vmwvvcinapha01.edificios.gfbanorte",
    "vmwvcsttapha01.edificios.gfbanorte"
)

# Log file path
$logFilePath = "$env:USERPROFILE\Desktop\SnapshotManagementLog.txt"

# Function to log messages
function Log-Message {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Add-Content -Path $logFilePath -Value $logEntry
    Write-Host $logEntry
}

# Function to select vCenter
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

# Connect to selected vCenter Server
$vCenterServer = Select-vCenter
$credential = Get-Credential
try {
    Connect-VIServer -Server $vCenterServer -Credential $credential -ErrorAction Stop
    Log-Message "Connected to vCenter: $vCenterServer"
} catch {
    Log-Message "Failed to connect to vCenter: $vCenterServer. Error: $_"
    exit
}

# Function to take snapshots
function Take-VMSnapshot {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$VMNames,
        [Parameter(Mandatory=$true)]
        [string]$SnapshotName,
        [Parameter(Mandatory=$true)]
        [string]$SnapshotDescription
    )

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

# Function to delete the last snapshot
function Remove-LastSnapshot {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$VMNames
    )

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

# Function to delete all snapshots
function Remove-AllSnapshots {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$VMNames
    )

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

# Function to delete selected snapshots
function Remove-SelectedSnapshots {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$VMNames
    )

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

# Function to restart the guest OS of VMs
function Restart-GuestOS {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$VMNames
    )

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

# Function to open file explorer and select CSV file
function Get-CSVFilePath {
    Add-Type -AssemblyName System.Windows.Forms
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop') # Default to Desktop
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

# Function to select VMs from a popup window
function Select-VMsFromList {
    $vms = Get-VM | Sort-Object Name
    $selectedVMs = $vms | Out-GridView -Title "Select VMs" -PassThru | Select-Object -ExpandProperty Name
    return $selectedVMs
}

# Main script logic
do {
    # Action selection
    while ($true) {
        Write-Host "Select an action to perform:"
        Write-Host "1. Take Snapshots"
        Write-Host "2. Delete the Last Snapshot"
        Write-Host "3. Delete All Snapshots"
        Write-Host "4. Delete Selected Snapshots"
        Write-Host "5. Restart Guest OS"
        $action = Read-Host "Enter your choice (1, 2, 3, 4, or 5)"
        if ($action -in 1..5) {
            break
        } else {
            Write-Host "Invalid choice. Please enter a number between 1 and 5."
        }
    }

    # VM selection
    while ($true) {
        Write-Host "Select an option to choose VMs:"
        Write-Host "1. Provide a CSV list with VM names"
        Write-Host "2. Select VMs from a popup window list"
        $option = Read-Host "Enter your choice (1 or 2)"
        if ($option -in 1..2) {
            break
        } else {
            Write-Host "Invalid choice. Please enter 1 or 2."
        }
    }

    if ($option -eq 1) {
        # Option 1: Provide a CSV list with VM names
        $csvPath = Get-CSVFilePath
        try {
            $vmNames = Import-Csv -Path $csvPath | Select-Object -ExpandProperty VMName -ErrorAction Stop
        } catch {
            Log-Message "Failed to read CSV file. Please ensure the file is valid and contains a 'VMName' column."
            continue
        }
    } elseif ($option -eq 2) {
        # Option 2: Select VMs from a popup window list
        $vmNames = Select-VMsFromList
    }

    if ($vmNames.Count -gt 0) {
        switch ($action) {
            1 {
                $snapshotName = Read-Host "Enter the snapshot name"
                $snapshotDescription = Read-Host "Enter the snapshot description"
                Take-VMSnapshot -VMNames $vmNames -SnapshotName $snapshotName -SnapshotDescription $snapshotDescription
            }
            2 {
                Remove-LastSnapshot -VMNames $vmNames
            }
            3 {
                Remove-AllSnapshots -VMNames $vmNames
            }
            4 {
                Remove-SelectedSnapshots -VMNames $vmNames
            }
            5 {
                Restart-GuestOS -VMNames $vmNames
            }
        }
    } else {
        Log-Message "No VMs selected."
    }

    # Ask if another action is needed
    while ($true) {
        $anotherAction = Read-Host "Do you want to perform another action? (yes/no)"
        if ($anotherAction -in "yes", "no") {
            break
        } else {
            Write-Host "Invalid input. Please enter 'yes' or 'no'."
        }
    }
} while ($anotherAction -eq "yes")

# Disconnect from vCenter Server
Disconnect-VIServer -Confirm:$false
Log-Message "Disconnected from vCenter: $vCenterServer"