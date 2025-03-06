<#
.SYNOPSIS
    Retrieves a list of Computer objects from a specified Active Directory OU, capturing logs in a transcript
    and optionally exporting results to CSV.

.DESCRIPTION
    Given an OU distinguished name (DN), this script:
      - Enumerates all computer objects in that OU and its sub-OUs.
      - Collects basic attributes for each Computer object (e.g., Name, OperatingSystem, etc.).
      - Automatically captures all output (including warnings/errors) in a transcript log file.
      - Suppresses warnings and errors in the console output, so they only appear in the transcript.
      - Optionally exports the results to CSV if -Output is provided.

.PARAMETER OU
    The distinguished name of the OU to search. Example: "OU=Workstations,DC=YourDomain,DC=com"

.PARAMETER Output
    Optional path to a CSV file. If provided, the script will export all Computer data to this file.

.PARAMETER TranscriptPath
    Optional. If provided, the script uses this path for the transcript log. Otherwise, it creates a
    timestamped file in the current directory.

.EXAMPLE
    .\Get-ADComputers.ps1 -OU "OU=Workstations,DC=YourDomain,DC=com"
    Retrieves all computers in that OU and sub-OUs, logs to a transcript file automatically.

.EXAMPLE
    .\Get-ADComputers.ps1 -OU "OU=Workstations,DC=YourDomain,DC=com" -Output "C:\Temp\WorkstationsList.csv"
    Retrieves all computers and exports them to CSV, while also logging to a transcript file.

.EXAMPLE
    .\Get-ADComputers.ps1 -OU "OU=Servers,DC=YourDomain,DC=com" -TranscriptPath "C:\Logs\ServerList.log"
    Logs all output to "C:\Logs\ServerList.log" instead of auto-generating a name, suppressing
    warnings/errors in the console but capturing them in the transcript.

.NOTES
    - Requires the Active Directory module (Import-Module ActiveDirectory).
    - PowerShell 5.1+ recommended.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OU,

    [Parameter(Mandatory = $false)]
    [string]$Output,

    [Parameter(Mandatory = $false)]
    [string]$TranscriptPath
)

# Store current preference settings so we can restore them later
$OldErrorPreference   = $ErrorActionPreference
$OldWarningPreference = $WarningPreference

# Suppress error messages and warning messages from being displayed in the console
$ErrorActionPreference = 'SilentlyContinue'
$WarningPreference     = 'SilentlyContinue'

try {
    # Ensure the AD module is loaded
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Error "Active Directory module not found or could not be loaded. Error: $($_.Exception.Message)"
    return
}

# If no transcript path was given, generate a default file name
if (-not $TranscriptPath) {
    $dateString     = (Get-Date -Format "yyyyMMdd-HHmmss")
    $TranscriptPath = ".\Get-ADComputers-$dateString.log"
}

# Attempt to start transcript logging
try {
    Start-Transcript -Path $TranscriptPath -Append | Out-Null
    Write-Verbose "Transcript logging started. Log file: $TranscriptPath"
} catch {
    Write-Warning "Failed to start transcript: $($_.Exception.Message)"
}

Write-Verbose "Querying AD for computers in OU: $OU..."

# Get all computer objects from the specified OU (and sub-OUs)
try {
    $adComputers = Get-ADComputer -Filter * -SearchBase $OU -SearchScope Subtree -ErrorAction Stop -Properties *
} catch {
    Write-Error "Failed to query AD. Make sure the OU is correct and you have permissions. Error: $($_.Exception.Message)"
    return
}

if (!$adComputers) {
    Write-Warning "No computer objects found in OU: $OU"
    return
}

Write-Verbose "Found $($adComputers.Count) computer(s). Gathering attributes..."

# Prepare a collection for results
$results = New-Object System.Collections.Generic.List[System.Object]

# Initialize progress bar
$totalComputers = $adComputers.Count
$progressCount  = 0

foreach ($adComputer in $adComputers) {
    $progressCount++
    $percentComplete = [math]::Round(($progressCount / $totalComputers) * 100, 2)

    Write-Progress -PercentComplete $percentComplete `
                   -Status "Retrieving attributes for $($adComputer.Name)" `
                   -Activity "Processing Computer $progressCount of $totalComputers"

    # Create a custom object with relevant AD attributes
    $computerObj = [PSCustomObject]@{
        Name                       = $adComputer.Name
        DistinguishedName          = $adComputer.DistinguishedName
        Description                = $adComputer.Description
        Location                   = $adComputer.Location
        OperatingSystem            = $adComputer.OperatingSystem
        OperatingSystemVersion     = $adComputer.OperatingSystemVersion
        OperatingSystemServicePack = $adComputer.OperatingSystemServicePack
        LastLogonDate             = $adComputer.LastLogonDate
        # Add more attributes as needed, or remove some you don't need
    }

    $results.Add($computerObj) | Out-Null
}

Write-Verbose "Attribute collection complete. $($results.Count) computers processed."

# If -Output was specified, export the results
if ($Output) {
    try {
        $results | Export-Csv -Path $Output -NoTypeInformation -Force
        Write-Information "Results exported to: $Output"
    } catch {
        Write-Warning "Failed to export to CSV. Error: $($_.Exception.Message)"
    }
} else {
    # Show results as a table
    Write-Host "`nResults:`n"
    $results | Format-Table
}

# Stop transcript (if started)
try {
    Stop-Transcript | Out-Null
    Write-Verbose "Transcript logging stopped. The log is saved at $TranscriptPath"
} catch {
    Write-Warning "Failed to stop transcript: $($_.Exception.Message)"
}

# Restore original preference settings
$ErrorActionPreference = $OldErrorPreference
$WarningPreference     = $OldWarningPreference

Write-Verbose "Done."
