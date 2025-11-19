# Get MAC Addresses from AD Computers (`Get-MACAddresses.ps1`)

This script queries Active Directory for computer objects in a given OU, attempts to contact each machine, and collects MAC addresses from the ones that are online.

---

## Features

- Queries an OU and its sub OUs for computer objects.
- Checks whether each computer is reachable.
- Attempts multiple retries for machines that are initially offline.
- Collects MAC addresses via CIM/WMI for online machines.
- Shows a progress bar and summary in the console.
- Can export results to CSV.
- Writes a transcript log containing warnings and errors.

---

## Requirements

- PowerShell 5.1 or later.
- Active Directory module:

  ```powershell
  Import-Module ActiveDirectory
  ````

* Permission to query the target computers with CIM or WMI.
* Network access from the machine running the script to the target computers.

---

## Parameters

* `-OU`
  Distinguished name of the OU to query. This is required. Example:

  ```powershell
  -OU "OU=Workstations,DC=example,DC=com"
  ```

* `-Output`
  Optional CSV file path where results will be saved.

* `-TranscriptPath`
  Optional path for the transcript log. If not supplied, a file such as `Get-MACAddresses-YYYYMMDD_HHMMSS.log` is created in the current directory.

---

## Usage

### Collect MAC addresses for workstations

```powershell
.\Get-MACAddresses.ps1 -OU "OU=Workstations,DC=example,DC=com"
```

The script will:

1. Pull a list of computers from the Workstations OU.
2. Attempt to contact each computer, retrying offline hosts a few times.
3. Gather MAC addresses for machines that respond.
4. Output a table with fields such as ComputerName, MacAddress, Status, and Error.
5. Write a transcript log.

### Export MAC addresses to CSV

```powershell
.\Get-MACAddresses.ps1 -OU "OU=Workstations,DC=example,DC=com" -Output "C:\Temp\WorkstationsMACs.csv"
```

This saves the results to `WorkstationsMACs.csv` for later use.

---

## Notes

* If your environment restricts remote CIM or WMI calls, you will need to ensure the necessary firewall rules and permissions are in place.
* The script can be extended to collect additional network information if you need more than just MAC addresses.