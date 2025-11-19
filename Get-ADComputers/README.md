# Get AD Computers (`Get-ADComputers.ps1`)

This script enumerates computer objects from a given OU in Active Directory and exports useful attributes such as name, operating system, description, and location.

---

## Features

- Queries an OU, including its sub OUs, for computer objects.
- Retrieves key attributes for reporting or inventory work.
- Shows a progress bar while it runs.
- Can export results to CSV.
- Creates a transcript log of the run.

---

## Requirements

- PowerShell 5.1 or later.
- Active Directory module:

  ```powershell
  Import-Module ActiveDirectory
  ```

* Read access to the OUs you are querying.

---

## Parameters

* `-OU`
  Distinguished name of the OU to query. This is required. Example:

  ```powershell
  -OU "OU=Servers,DC=example,DC=com"
  ```

* `-Output`
  Optional CSV file path for exported results.

* `-TranscriptPath`
  Optional path for the transcript log. If not supplied, a file such as `Get-ADComputers-YYYYMMDD_HHMMSS.log` is created in the current directory.

---

## Usage

### Basic inventory of all computers in an OU

```powershell
.\Get-ADComputers.ps1 -OU "OU=Servers,DC=example,DC=com"
```

This:

* Enumerates all computer objects under the `Servers` OU.
* Displays a table of results in the console.
* Writes a transcript file to the current directory.

### Export results to CSV

```powershell
.\Get-ADComputers.ps1 -OU "OU=Servers,DC=example,DC=com" -Output "C:\Temp\ServersList.csv"
```

This writes the results to `ServersList.csv` while still logging the run.

---

## Output

Typical fields in the output include:

* `Name`
* `OperatingSystem`
* `Description`
* `Location`
* Other attributes you select in the script

You can adjust the script to include additional properties if required.

---

## Notes

* For large environments, you may want to filter further by attributes such as operating system, name prefix, or enabled status.
* The script is a good starting point for feeding other automation that needs a list of computers.