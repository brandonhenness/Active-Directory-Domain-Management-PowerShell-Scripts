You are right, that README is out of date compared to this script.

Here is an updated `Get-ADComputers/README.md` that matches the current version you pasted, including the dynamic file name, `SearchBase` and `Filter` behavior, and the exported properties.

You can drop this straight into `Get-ADComputers/README.md`:

````md
# Get AD Computers (`Get-ADComputers.ps1`)

This script lists computer objects in Active Directory using a flexible `Get-ADComputer -Filter` expression. It writes results to a timestamped CSV file and logs the entire run to a transcript file next to the script.

The script is designed so that it does not hard code its own file name. The base name of the script file is used automatically in the log and CSV file names, so you can rename the script without breaking the logging behavior.

---

## Features

- Queries Active Directory for computer objects using any `-Filter` expression that `Get-ADComputer` supports.
- Optional `SearchBase` parameter to limit the search to a specific OU subtree.
- Defaults to searching the domain root (`defaultNamingContext`) when `SearchBase` is not supplied.
- Defaults to `-Filter *` when `Filter` is not supplied.
- Exports a detailed CSV with commonly useful properties.
- Records a transcript log for auditing and troubleshooting.
- Names the CSV and log files dynamically based on the script file name and a timestamp.

---

## Requirements

- Windows PowerShell 5.1 (Desktop edition).

  The script validates this and exits with an error if it is not running in Windows PowerShell 5.1 Desktop.

- Active Directory module:

  ```powershell
  Import-Module ActiveDirectory
````

* Sufficient permissions to query the domain and OUs you target.

---

## Parameters

### `-SearchBase`

Optional distinguished name (DN) of an OU to limit the search scope.

* If omitted, the script uses the domain root from `Get-ADRootDSE`:

  ```powershell
  (Get-ADRootDSE).defaultNamingContext
  ```

* If supplied, the script validates that the DN exists using `Get-ADObject`.

Example:

```powershell
-SearchBase "OU=Education,DC=example,DC=com"
```

### `-Filter`

Optional `Get-ADComputer -Filter` string.

* If omitted or empty, the script uses `*` which returns all computer objects in the search scope.
* The value must be a valid `-Filter` expression for `Get-ADComputer`.

Examples:

```powershell
-Filter 'Name -like "EDU-*"'
-Filter 'Enabled -eq $true'
-Filter 'OperatingSystem -like "Windows 11*" -and Enabled -eq $true'
```

---

## Output

The script creates two files in the same directory as the script:

* Transcript log:

  ```text
  <ScriptBaseName>_yyyyMMdd_HHmmss.log
  ```

* CSV export:

  ```text
  <ScriptBaseName>_yyyyMMdd_HHmmss.csv
  ```

`<ScriptBaseName>` is derived dynamically:

* If the script file is named `Get-ADComputers.ps1`, the files look like:

  ```text
  Get-ADComputers_20250101_130000.log
  Get-ADComputers_20250101_130000.csv
  ```

The CSV contains one row per computer with at least the following columns:

* `Name`
* `Description`
* `DNSHostName`
* `OperatingSystem`
* `OperatingSystemVersion`
* `IPv4Address`
* `Enabled`
* `LastLogonDate`
* `WhenCreated`
* `DistinguishedName`

The same information is also written to the console in a formatted table, which is captured in the transcript.

---

## Usage examples

### 1. List all computers in the domain

Searches the domain root (`defaultNamingContext`) and returns all computer objects.

```powershell
.\Get-ADComputers.ps1
```

### 2. Filter by operating system across the domain

Example for Windows 11 devices:

```powershell
.\Get-ADComputers.ps1 -Filter 'OperatingSystem -like "Windows 11*"'
```

### 3. List enabled computers in a specific OU subtree

```powershell
.\Get-ADComputers.ps1 `
    -SearchBase "OU=Servers,DC=example,DC=com" `
    -Filter 'Enabled -eq $true'
```

### 4. Filter by operating system within a specific OU subtree

```powershell
.\Get-ADComputers.ps1 `
    -SearchBase "OU=Workstations,DC=example,DC=com" `
    -Filter 'OperatingSystem -like "Windows 10*"'
```

After each run the script prints:

* The number of computers found.
* The path to the transcript log.
* The path to the CSV file.

---

## Notes

* If no matching computers are found, the script writes a message and exits without creating an empty CSV.
* The script wraps the main logic in a `try` block and writes a descriptive error if something goes wrong.
* Because the script validates the `SearchBase` DN, it will fail fast if you mistype an OU path instead of silently returning zero results.
* You can rename the script file without editing the code. The log and CSV file names follow the new script file name automatically.

```
::contentReference[oaicite:0]{index=0}
```
