# Create-EventLogFolders.ps1
# Script to create event log folder structure on multiple VMs
# Reads machine names from a CSV file
 
# Parameters
param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    [Parameter(Mandatory=$false)]
    [string]$MachineColumnName = "ComputerName",
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential
)
 
# Function to create folders on a remote machine
function Create-EventLogFolders {
    param (
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    Write-Host "Processing $ComputerName..." -ForegroundColor Cyan
    # Define remote script block
    $scriptBlock = {
        # Create base directory
        $baseDir = "D:\Logs"
        if (-not (Test-Path -Path $baseDir)) {
            New-Item -Path $baseDir -ItemType Directory -Force
        }
        # Create subdirectories
        $subDirs = @("Security", "System", "Application")
        foreach ($dir in $subDirs) {
            $fullPath = Join-Path -Path $baseDir -ChildPath $dir
            if (-not (Test-Path -Path $fullPath)) {
                New-Item -Path $fullPath -ItemType Directory -Force
                Write-Output "Created: $fullPath"
            } else {
                Write-Output "Already exists: $fullPath"
            }
        }
        # Check if D: drive exists
        if (-not (Test-Path -Path "D:\")) {
            Write-Output "WARNING: D: drive not found on this machine!"
            return $false
        }
        return $true
    }
    # Parameters for Invoke-Command
    $params = @{
        ComputerName = $ComputerName
        ScriptBlock = $scriptBlock
        ErrorAction = "Stop"
    }
    # Add credentials if provided
    if ($Credential) {
        $params.Add("Credential", $Credential)
    }
    try {
        $result = Invoke-Command @params
        if ($result) {
            Write-Host "Successfully created folder structure on $ComputerName" -ForegroundColor Green
        } else {
            Write-Host "Failed to create folder structure on $ComputerName - D: drive may not exist" -ForegroundColor Yellow
        }
        return $true
    } catch {
        Write-Host "Error creating folders on $ComputerName : $_" -ForegroundColor Red
        return $false
    }
}
 
# Main script execution
try {
    # Verify CSV file exists
    if (-not (Test-Path -Path $CsvPath)) {
        throw "CSV file not found at path: $CsvPath"
    }
    # Import CSV
    $machines = Import-Csv -Path $CsvPath
    # Verify machine column exists
    if (-not ($machines | Get-Member -Name $MachineColumnName -MemberType NoteProperty)) {
        throw "Column '$MachineColumnName' not found in CSV. Available columns: $($machines | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)"
    }
    # Summary counters
    $totalCount = $machines.Count
    $successCount = 0
    $failureCount = 0
    # Process each machine
    Write-Host "Starting folder creation on $totalCount machines..." -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor Cyan
    # Prompt for credentials if not provided
    if (-not $Credential) {
        $Credential = Get-Credential -Message "Enter credentials for remote computer access"
    }
    foreach ($machine in $machines) {
        $computerName = $machine.$MachineColumnName
        if ([string]::IsNullOrWhiteSpace($computerName)) {
            Write-Host "Skipping empty computer name in CSV" -ForegroundColor Yellow
            continue
        }
        $result = Create-EventLogFolders -ComputerName $computerName -Credential $Credential
        if ($result) {
            $successCount++
        } else {
            $failureCount++
        }
        Write-Host "----------------------------------------" -ForegroundColor Cyan
    }
    # Display summary
    Write-Host "Folder Creation Summary:" -ForegroundColor Cyan
    Write-Host "Total machines processed: $totalCount" -ForegroundColor White
    Write-Host "Successful: $successCount" -ForegroundColor Green
    Write-Host "Failed: $failureCount" -ForegroundColor Red
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
}