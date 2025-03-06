<#
.SYNOPSIS
    Retrieves MAC addresses for all computers in a specified Active Directory OU.
    Automatically logs all activity to a transcript file and suppresses warnings and errors in the console.

.DESCRIPTION
    Given an OU distinguished name (DN), this script:
      - Enumerates all computer objects in that OU and any sub-OUs.
      - Tests connectivity to each computer, allowing multiple retries.
      - Retrieves the MAC addresses via a CIM session if the computer is online.
      - Automatically captures all output (including warnings, errors, verbose messages) in a transcript.
      - Suppresses warning messages and error messages from displaying in the console.

.PARAMETER OU
    The distinguished name of the OU to search. Example:
    "OU=Workstations,DC=YourDomain,DC=com"

.PARAMETER Output
    Optional. If provided, the script will export results to this file.
    Example: "C:\Temp\ComputerMACs.csv"

.PARAMETER TranscriptPath
    Optional. If provided, the script uses this path for the transcript log.
    Otherwise, it creates a log file in the current directory with a timestamp.
    Example: "C:\Logs\MacAddressTranscript.log"

.EXAMPLE
    .\Get-MacAddresses.ps1 -OU "OU=Workstations,DC=YourDomain,DC=com"
    Retrieves MAC addresses for all computers under that OU,
    logs all activity to a default transcript file, and hides console errors and warnings.

.EXAMPLE
    .\Get-MacAddresses.ps1 -OU "OU=Servers,DC=YourDomain,DC=com" -Output "C:\Temp\ServersMACs.csv" -TranscriptPath "C:\Logs\MyTranscript.log"
    Retrieves MAC addresses, saves them to a CSV file, and logs to "MyTranscript.log" while suppressing console errors/warnings.

.NOTES
    Requires:
      - Windows PowerShell 5.1+ (or a PS version that supports AD and transcripts)
      - Active Directory module (Import-Module ActiveDirectory)
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
$OldErrorPreference    = $ErrorActionPreference
$OldWarningPreference  = $WarningPreference

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
    $dateString = (Get-Date -Format "yyyyMMdd-HHmmss")
    $TranscriptPath = ".\Get-MacAddresses-$dateString.log"
}

# Attempt to start transcript logging
try {
    Start-Transcript -Path $TranscriptPath -Append | Out-Null
    Write-Verbose "Transcript logging started. Log file: $TranscriptPath"
} catch {
    Write-Warning "Failed to start transcript: $($_.Exception.Message)"
}

Write-Verbose "Querying AD for computers in OU: $OU..."

# Get all computers from the specified OU (and sub-OUs)
try {
    $adComputers = Get-ADComputer -Filter * -SearchBase $OU -SearchScope Subtree -ErrorAction Stop
} catch {
    Write-Error "Failed to query AD. Make sure the OU is correct and you have permissions. Error: $($_.Exception.Message)"
    return
}

if (!$adComputers) {
    Write-Warning "No computers found in OU: $OU"
    return
}

Write-Verbose "Found $($adComputers.Count) computer(s). Beginning processing..."

# Prepare a collection for results
$results = New-Object System.Collections.Generic.List[System.Object]

# Initialize progress bar
$totalComputers = $adComputers.Count
$progressCount = 0

# Retry and delay parameters
$maxRetries = 3
$retryDelay = 5

foreach ($adComputer in $adComputers) {
    $computerName = $adComputer.Name
    $online = $false
    $retries = 0

    Write-Verbose "Checking connectivity for $computerName..."

    # Try to establish connection, retry if it fails
    while ($retries -lt $maxRetries -and -not $online) {
        $online = Test-Connection -ComputerName $computerName -Count 1 -Quiet
        if ($online) {
            Write-Verbose "$computerName is online."
        } else {
            Write-Verbose "$computerName is offline, retrying... ($($retries+1)/$maxRetries)"
            Start-Sleep -Seconds $retryDelay
            $retries++
        }
    }

    $macAddress = 'N/A'
    $status = 'Failed'
    $errorMsg = ''

    if ($online) {
        try {
            Write-Verbose "Attempting to retrieve MAC address from $computerName..."
            # Create a CIM session
            $options = New-CimSessionOption -UseSsl:$false -SkipCACheck -SkipCNCheck
            $session = New-CimSession -ComputerName $computerName -SessionOption $options

            # Retrieve network adapter information
            $netAdapter = Get-NetAdapter -CimSession $session

            # If multiple adapters exist, join them as a comma-separated string
            $macAddress = $netAdapter.MacAddress -join ', '
            $status = 'Success'

            # Close the session
            Remove-CimSession -CimSession $session

            Write-Information "MAC address retrieved for $($computerName): $macAddress"
        }
        catch {
            Write-Warning "Failed to retrieve data for $computerName. Error: $($_.Exception.Message)"
            $macAddress = 'N/A'
            $status = 'Failed'
            $errorMsg = $_.Exception.Message
        }
    } else {
        $errorMsg = 'Computer offline after retries.'
        Write-Verbose "$computerName did not come online after retries."
    }

    # Build a custom result object
    $result = [PSCustomObject]@{
        ComputerName = $computerName
        MacAddress   = $macAddress
        Status       = $status
        Error        = $errorMsg
    }

    # Add it to our results
    $results.Add($result)

    # Update progress
    $progressCount++
    $percentComplete = [math]::Round(($progressCount / $totalComputers) * 100, 2)
    Write-Progress -PercentComplete $percentComplete -Status "Processing $computerName" `
                   -Activity "Processing Computer $progressCount of $totalComputers"
}

Write-Verbose "Processing complete. Collected MAC addresses for $($results.Count) computer(s)."

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
