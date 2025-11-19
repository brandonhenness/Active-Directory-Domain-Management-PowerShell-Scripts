<#
.SYNOPSIS
    Creates Active Directory computer objects from a CSV input.

.DESCRIPTION
    This script reads a CSV file containing workstation metadata such as Site ID, Service Tag, Description,
    Organizational Unit, and optional AD groups. It automatically generates a computer name, creates the object
    in the specified OU, adds it to specified AD groups, and applies domain join permissions for a specified group.

    Logging is captured using Start-Transcript, and verbose output is supported. A CSV template can also be
    generated using the -GenerateTemplate switch.

.PARAMETER Path
    Optional path to the input CSV file. If not provided, a file selection dialog is shown.

.PARAMETER TranscriptPath
    Optional path to save the transcript log. Defaults to a timestamped file in the current directory.

.PARAMETER GenerateTemplate
    If specified, creates a sample CSV template and exits.

.PARAMETER RetryCount
    Times to retry looking up the computer after creation before giving up. Default 10.

.PARAMETER RetryDelaySeconds
    Seconds to wait between retries. Default 3.

.EXAMPLE
    .\Create-ADComputers.ps1 -Path "computers.csv" -TranscriptPath "log.txt" -Verbose

.EXAMPLE
    .\Create-ADComputers.ps1 -GenerateTemplate
    Creates a template file named ComputerTemplate.csv.

.NOTES
    Author: Brandon Henness
    Updated: 2025-11-05
    Requires: ActiveDirectory module, PowerShell 5.1+
#>

[CmdletBinding()]
param(
    [string]$Path,
    [string]$TranscriptPath,
    [switch]$GenerateTemplate,
    [int]$RetryCount = 10,
    [int]$RetryDelaySeconds = 3
)

# PowerShell version check
if (-not ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -eq 1)) {
    Write-Host "This script must be run in Windows PowerShell 5.1 (not PowerShell 7+)." -ForegroundColor Red
    Write-Host "Detected version: $($PSVersionTable.PSVersion.ToString())" -ForegroundColor Yellow
    Write-Host "`nTo open PowerShell 5.1, run 'powershell.exe' instead of 'pwsh.exe'." -ForegroundColor Cyan
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms

# Backup current error and warning preferences
$OldErrorPreference    = $ErrorActionPreference
$OldWarningPreference  = $WarningPreference

# Suppress errors and warnings in console
$ErrorActionPreference = 'SilentlyContinue'
$WarningPreference     = 'SilentlyContinue'

# Start transcript logging
if (-not $TranscriptPath) {
    $dateString = (Get-Date -Format "yyyyMMdd-HHmmss")
    $TranscriptPath = ".\Create-ADComputers-$dateString.log"
}

try {
    Start-Transcript -Path $TranscriptPath -Append | Out-Null
} catch {
    Write-Warning "Failed to start transcript: $($_.Exception.Message)"
}

# Generate template CSV and exit if requested
if ($GenerateTemplate) {
    $template = @( [PSCustomObject]@{
        'Site ID'     = 'SITE01'
        'Service Tag' = 'ABC1234'
        'Description' = 'Example Workstation'
        'OU'          = 'OU=Workstations,DC=domain,DC=com'
        'Group1'      = 'Workstation Group'
        'Group2'      = ''
        'Group3'      = ''
        'Group4'      = ''
        'Group5'      = ''
        'JoinGroup'   = 'OSN\ElevatedComputerJoiners'
    })
    $template | Export-Csv -Path ".\ComputerTemplate.csv" -NoTypeInformation -Force
    Write-Host "Template CSV created at .\ComputerTemplate.csv" -ForegroundColor Green
    try { Stop-Transcript | Out-Null } catch {}
    return
}

# File selection dialog if Path is not provided
function Select-CsvFile {
    [System.Windows.Forms.OpenFileDialog]$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $openFileDialog.Title = "Select the CSV File"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    } else {
        Write-Warning "No file selected. Exiting..."
        try { Stop-Transcript | Out-Null } catch {}
        exit
    }
}

if (-not $Path) {
    $Path = Select-CsvFile
}

Write-Verbose "Selected CSV file: $Path"

try {
    $computers = Import-Csv -Path $Path -ErrorAction Stop
    Write-Verbose "Successfully imported CSV data."
} catch {
    Write-Error "Failed to import CSV file. Error: $($_.Exception.Message)"
    try { Stop-Transcript | Out-Null } catch {}
    exit
}

function New-ComputerName {
    param (
        [string]$SiteID,
        [string]$ServiceTag
    )
    $SiteID = $SiteID.Trim()
    $ServiceTag = $ServiceTag.Trim()

    $domainPrefix = "OSN"
    $maxLength = 15
    $availableLength = $maxLength - ($domainPrefix.Length + $SiteID.Length)

    if ($availableLength -gt 0) {
        $truncatedTag = $ServiceTag.Substring([Math]::Max(0, $ServiceTag.Length - $availableLength))
    } else {
        Write-Error "SiteID '$SiteID' is too long to fit in computer name."
        throw "Invalid SiteID length for computer name generation."
    }

    return "$domainPrefix$SiteID$truncatedTag"
}

function New-ADComputerObject {
    param (
        [string]$Name,
        [string]$Description,
        [string]$OU
    )
    $Name = $Name.Trim()
    $Description = $Description.Trim()
    $OU = $OU.Trim()

    try {
        New-ADComputer -Name $Name -Path $OU -Description $Description -PassThru -ErrorAction Stop | Out-Null
        Write-Verbose "Computer object '$Name' created in OU '$OU'."
    } catch {
        Write-Error "Failed to create computer object '$Name'. Error: $($_.Exception.Message)"
    }
}

function Add-ComputerToGroups {
    param (
        [string]$ComputerName,
        [string[]]$Groups
    )
    $ComputerName = $ComputerName.Trim()
    foreach ($group in $Groups) {
        if (![string]::IsNullOrWhiteSpace($group)) {
            $grp = $group.Trim()
            try {
                Add-ADGroupMember -Identity $grp -Members "$ComputerName$" -ErrorAction Stop
                Write-Verbose "Computer '$ComputerName' added to group '$grp'."
            } catch {
                Write-Warning "Failed to add '$ComputerName' to '$grp'. Error: $($_.Exception.Message)"
            }
        }
    }
}

function Get-ADComputerRetry {
    <#
    .SYNOPSIS
        Looks up a computer by sAMAccountName with retries to ride out replication.
    .DESCRIPTION
        Uses -Filter on sAMAccountName to avoid DN vs Name ambiguity and to match the
        trailing $ used by computer accounts.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [int]$RetryCount = 10,
        [int]$RetryDelaySeconds = 3
    )

    $ComputerName = $ComputerName.Trim()
    $sam = "$ComputerName$"
    for ($i = 0; $i -lt $RetryCount; $i++) {
        try {
            $obj = Get-ADComputer -Filter "sAMAccountName -eq '$sam'" -Properties DistinguishedName -ErrorAction Stop
            if ($obj) { return $obj }
        } catch {
            # swallow and retry
        }
        Start-Sleep -Seconds $RetryDelaySeconds
    }
    return $null
}

function Set-JoinDomainPermission {
    <#
    .SYNOPSIS
        Sets domain join permissions for a group on a computer object.
    .DESCRIPTION
        Grants a group full control on the specific computer object with no inheritance
        by modifying the object's ACL using the AD drive. Uses retry to ensure the object
        is visible before ACL work.
    #>
    param (
        [string]$ComputerName,
        [string]$GroupName,
        [int]$RetryCount,
        [int]$RetryDelaySeconds
    )

    if ([string]::IsNullOrWhiteSpace($GroupName)) {
        Write-Verbose "No JoinGroup provided for $ComputerName. Skipping ACL."
        return
    }

    # Confirm the computer exists and we can resolve its DN
    $computer = Get-ADComputerRetry -ComputerName $ComputerName -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds
    if (-not $computer) {
        Write-Warning "Computer '$ComputerName' not found after $RetryCount attempts. Skipping ACL."
        return
    }

    $dn = $computer.DistinguishedName
    $ADPath = "AD:\$dn"
    Write-Verbose "Setting permissions on AD path: $ADPath"

    # Get current ACL
    try {
        $acl = Get-Acl -Path $ADPath -ErrorAction Stop
    } catch {
        Write-Warning "Failed to read ACL for '$ComputerName' at $ADPath. Error: $($_.Exception.Message)"
        return
    }

    # Resolve the target group SID
    try {
        $ntAccount = New-Object System.Security.Principal.NTAccount($GroupName.Trim())
        $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
    } catch {
        Write-Warning "Could not resolve group '$GroupName' to SID. Skipping ACL update for '$ComputerName'. Error: $($_.Exception.Message)"
        return
    }

    # Avoid duplicate rule for the same SID with GenericAll on the object
    $hasRule = $false
    foreach ($rule in $acl.Access) {
        try {
            $ruleSid = $rule.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier])
            if ($ruleSid -eq $sid -and ($rule.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::GenericAll)) {
                $hasRule = $true
                break
            }
        } catch {
            # ignore translation issues and continue
        }
    }
    if ($hasRule) {
        Write-Verbose "ACL already grants GenericAll to '$GroupName' on '$ComputerName'. Skipping add."
        return
    }

    # Build access rule for object only (no inheritance)
    $rights      = [System.DirectoryServices.ActiveDirectoryRights]::GenericAll
    $controlType = [System.Security.AccessControl.AccessControlType]::Allow
    $inheritance = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::None

    try {
        $adRule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($sid, $rights, $controlType, $inheritance)
        $acl.AddAccessRule($adRule)
        Set-Acl -Path $ADPath -AclObject $acl -ErrorAction Stop
        Write-Verbose "Successfully granted full control to '$GroupName' on '$ComputerName'."
    } catch {
        Write-Warning "Failed to update ACL on '$ComputerName'. Error: $($_.Exception.Message)"
    }
}

Write-Verbose "Script execution started."

$total = $computers.Count
$count = 0

foreach ($computer in $computers) {
    $count++

    # Normalize and trim CSV inputs
    $siteID      = ($computer."Site ID"     | ForEach-Object { $_.ToString().Trim() })
    $serviceTag  = ($computer."Service Tag" | ForEach-Object { $_.ToString().Trim() })
    $description = ($computer.Description   | ForEach-Object { $_.ToString().Trim() })
    $ouLocation  = ($computer.OU            | ForEach-Object { $_.ToString().Trim() })
    $groups      = @($computer.Group1, $computer.Group2, $computer.Group3, $computer.Group4, $computer.Group5) | ForEach-Object { if ($_ -ne $null) { $_.ToString().Trim() } }
    $joinGroup   = ($computer.JoinGroup     | ForEach-Object { $_.ToString().Trim() })

    try {
        $computerName = New-ComputerName -SiteID $siteID -ServiceTag $serviceTag
        Write-Verbose "Generated computer name: $computerName"
    } catch {
        Write-Warning "Name generation failed for Site ID '$siteID'. Skipping."
        continue
    }

    # Create object
    New-ADComputerObject -Name $computerName -Description $description -OU $ouLocation

    # Add to groups
    Add-ComputerToGroups -ComputerName $computerName -Groups $groups

    # Apply join permissions with retry
    Set-JoinDomainPermission -ComputerName $computerName -GroupName $joinGroup -RetryCount $RetryCount -RetryDelaySeconds $RetryDelaySeconds

    Write-Progress -Activity "Creating AD Computers" -Status "Processing $computerName" -PercentComplete (($count / $total) * 100)
}

Write-Verbose "Script execution completed."
Write-Host "All computer objects processed." -ForegroundColor Green

try { Stop-Transcript | Out-Null } catch { Write-Warning "Failed to stop transcript." }

# Restore original preferences
$ErrorActionPreference = $OldErrorPreference
$WarningPreference     = $OldWarningPreference
