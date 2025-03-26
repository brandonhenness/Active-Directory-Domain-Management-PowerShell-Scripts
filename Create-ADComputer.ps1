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

.EXAMPLE
    .\Create-ADComputers.ps1 -Path "computers.csv" -TranscriptPath "log.txt" -Verbose

.EXAMPLE
    .\Create-ADComputers.ps1 -GenerateTemplate
    Creates a template file named ComputerTemplate.csv.

.NOTES
    Author: Brandon Henness
    Updated: 2025-03-21
    Requires: ActiveDirectory module, PowerShell 5.1+
#>

[CmdletBinding()]
param(
    [string]$Path,
    [string]$TranscriptPath,
    [switch]$GenerateTemplate
)

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
    Stop-Transcript | Out-Null
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
        Stop-Transcript | Out-Null
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
    Stop-Transcript | Out-Null
    exit
}

function New-ComputerName {
    param (
        [string]$SiteID,
        [string]$ServiceTag
    )
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
    foreach ($group in $Groups) {
        if (![string]::IsNullOrWhiteSpace($group)) {
            try {
                Add-ADGroupMember -Identity $group -Members "$ComputerName$" -ErrorAction Stop
                Write-Verbose "Computer '$ComputerName' added to group '$group'."
            } catch {
                Write-Warning "Failed to add '$ComputerName' to '$group'. Error: $($_.Exception.Message)"
            }
        }
    }
}

function Set-JoinDomainPermission {
    <#
    .SYNOPSIS
        Sets domain join permissions for a group on a computer object.
    .DESCRIPTION
        Grants a group full control on the specific computer object (no inheritance)
        by modifying the object's ACL using the AD: drive (Get-Acl/Set-Acl).
    #>
    param (
        [string]$ComputerName,
        [string]$GroupName
    )
    # Retrieve the computer object's DN using Get-ADComputer.
    $computer = Get-ADComputer -Identity $ComputerName -Properties DistinguishedName
    if (-not $computer) {
        throw "Computer '$ComputerName' not found."
    }
    $dn = $computer.DistinguishedName
    $ADPath = "AD:\$dn"
    Write-Verbose "Setting permissions on AD path: $ADPath"
    
    # Retrieve the current ACL.
    $acl = Get-Acl -Path $ADPath
    
    # Get the group's SID.
    $ntAccount = New-Object System.Security.Principal.NTAccount($GroupName)
    $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier])
    
    # Define full control rights using GenericAll.
    $rights = [System.DirectoryServices.ActiveDirectoryRights]::GenericAll
    $controlType = [System.Security.AccessControl.AccessControlType]::Allow
    $inheritance = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::None
    
    # Create and add the access rule.
    $accessRule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule (
        $sid, $rights, $controlType, $inheritance
    )
    $acl.AddAccessRule($accessRule)
    
    # Write the updated ACL back to the computer object.
    Set-Acl -Path $ADPath -AclObject $acl -ErrorAction Stop
    Write-Verbose "Successfully granted full control to '$GroupName' on '$ComputerName'"
}

Write-Verbose "Script execution started."

$total = $computers.Count
$count = 0

foreach ($computer in $computers) {
    $count++
    $siteID = $computer."Site ID"
    $serviceTag = $computer."Service Tag"
    $description = $computer.Description
    $ouLocation = $computer.OU
    $groups = @($computer.Group1, $computer.Group2, $computer.Group3, $computer.Group4, $computer.Group5)
    $joinGroup = $computer.JoinGroup

    try {
        $computerName = New-ComputerName -SiteID $siteID -ServiceTag $serviceTag
        Write-Verbose "Generated computer name: $computerName"
    } catch {
        Write-Warning "Name generation failed for Site ID '$siteID'. Skipping."
        continue
    }

    New-ADComputerObject -Name $computerName -Description $description -OU $ouLocation
    Add-ComputerToGroups -ComputerName $computerName -Groups $groups

    # Apply join permissions by granting full control to the join group.
    Set-JoinDomainPermission -ComputerName $computerName -GroupName $joinGroup

    Write-Progress -Activity "Creating AD Computers" -Status "Processing $computerName" -PercentComplete (($count / $total) * 100)
}

Write-Verbose "Script execution completed."
Write-Host "All computer objects processed." -ForegroundColor Green

try { Stop-Transcript | Out-Null } catch { Write-Warning "Failed to stop transcript." }

# Restore original preferences
$ErrorActionPreference = $OldErrorPreference
$WarningPreference     = $OldWarningPreference
