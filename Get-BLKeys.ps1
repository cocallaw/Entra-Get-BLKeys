Param (
    [string]$MachinesCSV
)

# Validate that the file exists and has a .csv extension
if (-not (Test-Path -Path $MachinesCSV)) {
    Write-Error "The file '$MachinesCSV' does not exist. Please provide a valid file path."
    exit
}

if (-not ($MachinesCSV -like "*.csv")) {
    Write-Error "The file '$MachinesCSV' is not a CSV file. Please provide a valid CSV file."
    exit
}

# Validate that the CSV file contains the required 'VMName' column
$csvContent = Import-Csv -Path $MachinesCSV
if (-not $csvContent[0].PSObject.Properties.Name -contains "MachineName") {
    Write-Error "The CSV file does not contain the required 'Machine' column. Please provide a valid CSV file."
    exit
}

# Check if Microsoft Graph PowerShell module is installed and up-to-date
$module = Get-Module -ListAvailable -Name Microsoft.Graph
if (-not $module) {
    Write-Error "Microsoft Graph PowerShell module is not installed. Please install it using 'Install-Module Microsoft.Graph -Scope CurrentUser'."
    exit
} elseif ($module.Version -lt [Version]"2.5.0") {
    Write-Error "The installed Microsoft Graph module is outdated. Please update it using 'Update-Module Microsoft.Graph'."
    exit
}

# Connect to Microsoft Graph with necessary scopes
try {
    Connect-MgGraph -Scopes "Device.Read.All", "BitLockerKey.Read.All" -NoWelcome
} catch {
    Write-Error "Failed to connect to Microsoft Graph. Please ensure you have the necessary permissions and try again."
    exit
}

# Import the Microsoft Graph Identity Directory Management module
Import-Module Microsoft.Graph.Identity.DirectoryManagement -Force

# Read the CSV file and extract VM names
$vmNames = (Import-Csv -Path $MachinesCSV).MachineName

$devices = @()
foreach ($vmName in $vmNames) {
    $devices += Get-MgDevice -Filter "displayName eq '$vmName'"
}

# Initialize an array to store device recovery keys
$deviceRecoveryKeys = @()

# Loop through each device and retrieve its BitLocker recovery keys
foreach ($device in $devices) {
    $recoveryKeys = Get-MgInformationProtectionBitlockerRecoveryKey -Filter "deviceId eq '$($device.DeviceId)'"
    if ($recoveryKeys) {
        foreach ($key in $recoveryKeys) {
            $deviceRecoveryKeys += [PSCustomObject]@{
                DeviceName    = $device.DisplayName
                DeviceId      = $device.DeviceId
                RecoveryKeyId = $key.Id
                RecoveryKey   = (Get-MgInformationProtectionBitlockerRecoveryKey -BitlockerRecoveryKeyId $key.Id -Property key).key
            }
        }
    } else {
        # If no recovery keys are found, set default values
        $deviceRecoveryKeys += [PSCustomObject]@{
            DeviceName    = $device.DisplayName
            DeviceId      = $device.DeviceId
            RecoveryKeyId = "NoKeyFound"
            RecoveryKey   = "NoKeyFound"
        }
    }
}

# Export the device recovery keys to a CSV file
$outputFilePath = "DeviceRecoveryKeys.csv"
$deviceRecoveryKeys | Export-Csv -Path $outputFilePath -NoTypeInformation

# Get the absolute path of the output file
$absolutePath = Resolve-Path -Path $outputFilePath

# Display the full file path in the output message
Write-Output "Device recovery keys have been exported to '$absolutePath'."