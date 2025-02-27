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

# Function to select vCenter
function Select-vCenter {
    Write-Host "Select a vCenter to connect to:"
    for ($i = 0; $i -lt $vCenters.Length; $i++) {
        Write-Host "$($i + 1). $($vCenters[$i])"
    }
    $selection = Read-Host "Enter the number of the vCenter"
    if ($selection -ge 1 -and $selection -le $vCenters.Length) {
        return $vCenters[$selection - 1]
    } else {
        Write-Host "Invalid selection. Exiting script."
        exit
    }
}

# Connect to selected vCenter Server
$vCenterServer = Select-vCenter
$credential = Get-Credential
Connect-VIServer -Server $vCenterServer -Credential $credential

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
        $vm = Get-VM -Name $vmName -ErrorAction SilentlyContinue
        if ($vm) {
            New-Snapshot -VM $vm -Name $SnapshotName -Description $SnapshotDescription -Confirm:$false
            Write-Host "Snapshot taken for VM: $vmName"
        } else {
            Write-Host "VM not found: $vmName"
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
        $vm = Get-VM -Name $vmName -ErrorAction SilentlyContinue
        if ($vm) {
            $lastSnapshot = Get-Snapshot -VM $vm | Sort-Object Created -Descending | Select-Object -First 1
            if ($lastSnapshot) {
                Remove-Snapshot -Snapshot $lastSnapshot -Confirm:$false
                Write-Host "Last snapshot deleted for VM: $vmName"
            } else {
                Write-Host "No snapshots found for VM: $vmName"
            }
        } else {
            Write-Host "VM not found: $vmName"
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
        $vm = Get-VM -Name $vmName -ErrorAction SilentlyContinue
        if ($vm) {
            $snapshots = Get-Snapshot -VM $vm
            if ($snapshots) {
                Remove-Snapshot -Snapshot $snapshots -Confirm:$false
                Write-Host "All snapshots deleted for VM: $vmName"
            } else {
                Write-Host "No snapshots found for VM: $vmName"
            }
        } else {
            Write-Host "VM not found: $vmName"
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
        Write-Host "No file selected. Exiting script."
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
Write-Host "Select an action to perform:"
Write-Host "1. Take Snapshots"
Write-Host "2. Delete the Last Snapshot"
Write-Host "3. Delete All Snapshots"
$action = Read-Host "Enter your choice (1, 2, or 3)"

Write-Host "Select an option to choose VMs:"
Write-Host "1. Provide a CSV list with VM names"
Write-Host "2. Select VMs from a popup window list"
$option = Read-Host "Enter your choice (1 or 2)"

if ($option -eq 1) {
    # Option 1: Provide a CSV list with VM names
    $csvPath = Get-CSVFilePath
    $vmNames = Import-Csv -Path $csvPath | Select-Object -ExpandProperty VMName
} elseif ($option -eq 2) {
    # Option 2: Select VMs from a popup window list
    $vmNames = Select-VMsFromList
} else {
    Write-Host "Invalid option selected. Exiting script."
    Disconnect-VIServer -Confirm:$false
    exit
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
        default {
            Write-Host "Invalid action selected. Exiting script."
        }
    }
} else {
    Write-Host "No VMs selected. Exiting script."
}

# Disconnect from vCenter Server
Disconnect-VIServer -Confirm:$false