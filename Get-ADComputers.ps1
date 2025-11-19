<#
.SYNOPSIS
  List computers in Active Directory. Logs to transcript and exports full details to CSV.

.DESCRIPTION
  Queries Active Directory for computer objects using a flexible Get-ADComputer -Filter
  expression so you can filter on any computer attribute. Results are written to a CSV
  and a transcript log next to the script.

.PARAMETER SearchBase
  Optional distinguished name of an OU to limit the search scope.
  If omitted, the script searches the domain root (defaultNamingContext).

  Example:
    -SearchBase "OU=Education,DC=osn,DC=wa,DC=gov"

.PARAMETER Filter
  Optional Get-ADComputer -Filter string.
  If omitted, the script uses -Filter * (all computers).

  Examples:
    -Filter 'Name -like "EDU-*"'
    -Filter 'Enabled -eq $true'
    -Filter 'OperatingSystem -like "Windows 11*" -and Enabled -eq $true'

.NOTES
  - Requires the Active Directory module (Import-Module ActiveDirectory).
  - Requires Windows PowerShell 5.1 (Desktop edition).

.EXAMPLE
  .\ScriptName.ps1
  Searches the entire domain and returns all computer objects.

.EXAMPLE
  .\ScriptName.ps1 -Filter 'OperatingSystem -like "Windows 11*"'
  Searches the entire domain for computers where OperatingSystem matches Windows 11.

.EXAMPLE
  .\ScriptName.ps1 -SearchBase "OU=SBCTC,OU=Education,DC=osn,DC=wa,DC=gov" -Filter 'Enabled -eq $true'
  Searches a specific OU subtree for enabled computers only.

.EXAMPLE
  .\ScriptName.ps1 -SearchBase "OU=Education,DC=osn,DC=wa,DC=gov" -Filter 'OperatingSystem -like "Windows 10*"'
  Searches a specific OU subtree for computers where OperatingSystem matches Windows 10.
#>

[CmdletBinding()]
param(
    [Parameter(
        Mandatory = $false,
        HelpMessage = 'Optional distinguished name of an OU to limit the search scope.'
    )]
    [string]$SearchBase,

    [Parameter(
        Mandatory = $false,
        HelpMessage = 'Get-ADComputer -Filter string (e.g. ''Enabled -eq $true''). Defaults to *.'
    )]
    [string]$Filter
)

# Require Windows PowerShell 5.1 Desktop
if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error 'Please run this script in Windows PowerShell 5.1 (Desktop edition).'
    return
}

# Import AD module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Error "Failed to import ActiveDirectory module. $($_.Exception.Message)"
    return
}

# Build paths next to the script
$scriptDir  = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

# Derive the script file name (without extension) dynamically
if ($MyInvocation.MyCommand.Path) {
    $scriptFileName = Split-Path -Leaf $MyInvocation.MyCommand.Path
    $scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($scriptFileName)
}
else {
    # Fallback if run in an interactive host with no script path
    $scriptBaseName = 'ADComputerQuery'
}

$timestamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
$transcriptPath = Join-Path $scriptDir ("{0}_{1}.log" -f $scriptBaseName, $timestamp)
$csvPath        = Join-Path $scriptDir ("{0}_{1}.csv" -f $scriptBaseName, $timestamp)

Start-Transcript -Path $transcriptPath -Append | Out-Null

try {
    # Determine search base
    if ([string]::IsNullOrWhiteSpace($SearchBase)) {
        $SearchBase = (Get-ADRootDSE -ErrorAction Stop).defaultNamingContext
        Write-Host "Search base not provided. Using domain root: $SearchBase"
    }
    else {
        # Validate the provided DN exists
        $null = Get-ADObject -Identity $SearchBase -ErrorAction Stop
        Write-Host "Using provided search base: $SearchBase"
    }

    # Build AD filter
    if ([string]::IsNullOrWhiteSpace($Filter)) {
        $Filter = '*'
    }

    Write-Host "Using AD filter: $Filter"
    Write-Host "Querying for computer objects..."

    $computers = Get-ADComputer `
        -SearchBase $SearchBase `
        -SearchScope Subtree `
        -Filter $Filter `
        -Properties Description, OperatingSystem, OperatingSystemVersion, IPv4Address, DNSHostName, LastLogonDate, whenCreated, Enabled

    if (-not $computers) {
        Write-Host 'No matching computers found.'
        return
    }

    # Shape results using PSCustomObject and sort
    $results = $computers |
        ForEach-Object {
            [PSCustomObject]@{
                Name                   = $_.Name
                Description            = $_.Description
                DNSHostName            = $_.DNSHostName
                OperatingSystem        = $_.OperatingSystem
                OperatingSystemVersion = $_.OperatingSystemVersion
                IPv4Address            = $_.IPv4Address
                Enabled                = $_.Enabled
                LastLogonDate          = $_.LastLogonDate
                WhenCreated            = $_.whenCreated
                DistinguishedName      = $_.DistinguishedName
            }
        } |
        Sort-Object Name

    # Write summary to console (captured in transcript)
    Write-Host ''
    Write-Host "Found $($results.Count) computer(s)."
    Write-Host "Saving full results to CSV: $csvPath"
    Write-Host ''

    $results | Format-Table -AutoSize
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host ''
    Write-Host "Transcript saved to: $transcriptPath"
    Write-Host "CSV saved to:        $csvPath"
}
catch {
    Write-Error "An error occurred. $($_.Exception.Message)"
}
finally {
    Stop-Transcript | Out-Null
}
