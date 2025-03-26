# Active Directory & Domain Management PowerShell Scripts

A collection of **PowerShell scripts** designed to help with day-to-day **Active Directory** (AD) tasks and the management of **domain-joined computers**. Each script focuses on a different aspect of administration, such as user management, computer inventory, network configuration, and more.

## Table of Contents

1. [Overview](#overview)  
2. [Scripts Included](#scripts-included)  
3. [Requirements](#requirements)  
4. [Usage](#usage)  
5. [Configuration & Customization](#configuration--customization)  
6. [Contributing](#contributing)  
7. [License](#license)

---

## Overview

Managing an Active Directory environment can be time-consuming when done manually. By using these PowerShell scripts, you can automate many of the common tasks involved in maintaining your domain and its computers. Examples of tasks covered include:

- Collecting hardware or network information (MAC addresses, IPs, etc.).  
- Retrieving or updating user and computer objects in AD.  
- Inventorying domain-joined machines, performing connectivity checks, or mass updates.  
- **Creating computer objects from CSV input while automatically applying domain join permissions.**

Scripts in this repo can be used independently or chained together, depending on your workflow.

---

## Scripts Included

### 1. **`Get-MacAddresses.ps1`**  
**Purpose**: Queries an OU (and any sub-OUs) in Active Directory for computer objects, checks which are online, and retrieves their MAC addresses.  
- **Key Features**:
  - **Requires** an **OU** distinguished name as a parameter (e.g., `-OU "OU=Workstations,DC=YourDomain,DC=com"`).  
  - **Retries** offline machines up to three times before marking them as offline.  
  - **Suppresses** console warnings/errors (but logs them in a transcript).  
  - **Outputs** a dynamic progress bar and final table of results.  
  - **Optional** CSV export (via the `-Output` parameter).  
  - **Automatic transcript logging** if no `-TranscriptPath` is provided.

### 2. **`Get-ADComputers.ps1`**  
**Purpose**: Enumerates all **Computer** objects from a specified OU (and any sub-OUs), retrieving key attributes like **Name**, **Description**, **Location**, **OS**, etc.  
- **Key Features**:
  - Accepts an **OU** parameter (e.g., `-OU "OU=Servers,DC=YourDomain,DC=com"`).  
  - **Suppresses** console warnings/errors, logging them in a transcript.  
  - **Displays** a progress bar as it enumerates AD computers.  
  - **Allows** CSV export (via the `-Output` parameter).  
  - **Automatic transcript logging** if no `-TranscriptPath` is provided.

### 3. **`Create-ADComputers.ps1`**  
**Purpose**: Creates AD computer objects from a CSV file containing workstation metadata and immediately applies join permissions to the objects.  
- **Key Features**:
  - **CSV Input**: Reads a CSV file containing details such as Site ID, Service Tag, Description, OU, and various AD groups.  
  - **Dynamic Name Generation**: Automatically generates a computer name based on the Site ID and Service Tag.  
  - **Object Creation**: Creates the computer object in the specified OU.  
  - **Group Assignment**: Adds the new computer to any specified AD groups.  
  - **Join Permission Application**: Automatically applies domain join permissions by granting full control (GenericAll) to a specified join group. This is done via an ADSI binding with a triple-slash LDAP URL to ensure the correct DC is targeted.
  - **Logging & Transcript**: Captures a transcript log (or uses a specified path) and provides verbose output.  
  - **Template Generation**: Can generate a sample CSV template using the `-GenerateTemplate` switch.

---

## Requirements

1. **PowerShell 5.1+**  
2. **Active Directory module** (for commands like `Get-ADUser`, `Get-ADComputer`, etc.):  
   ```powershell
   Import-Module ActiveDirectory
   ```
3. **Appropriate permissions** within your domain environment.  
4. **Network firewall rules** allowing connectivity to/from domain controllers and client machines, if needed.

---

## Usage

### `Get-MacAddresses.ps1`

**Example:**  
```powershell
.\Get-MacAddresses.ps1 -OU "OU=Workstations,DC=YourDomain,DC=com"
```
- Automatically creates a transcript log named `Get-MacAddresses-<timestamp>.log` in the current directory.  
- Suppresses console warnings/errors but captures them in the log.  
- Outputs a table of **ComputerName**, **MacAddress**, **Status**, and any **Error** messages.

To **export** results to CSV:  
```powershell
.\Get-MacAddresses.ps1 -OU "OU=Workstations,DC=YourDomain,DC=com" -Output "C:\Temp\WorkstationsMACs.csv"
```

### `Get-ADComputers.ps1`

**Example:**  
```powershell
.\Get-ADComputers.ps1 -OU "OU=Servers,DC=YourDomain,DC=com"
```
- Retrieves all AD computer objects under the specified **Servers** OU.  
- Logs all activities and messages in a transcript file named `Get-ADComputers-<timestamp>.log`.  
- Displays a final table of objects, including **Name**, **OperatingSystem**, **Description**, **Location**, etc.

To **export** results to CSV:  
```powershell
.\Get-ADComputers.ps1 -OU "OU=Servers,DC=YourDomain,DC=com" -Output "C:\Temp\ServersList.csv"
```

### `Create-ADComputers.ps1`

**Example:**  
```powershell
.\Create-ADComputers.ps1 -CsvPath "computers.csv" -TranscriptPath "log.txt" -Verbose
```
- **CSV Input**: The script reads the provided CSV file (or opens a file selection dialog if none is provided).  
- **Computer Creation**: It generates a computer name, creates the computer object in the specified OU, and adds it to any listed AD groups.  
- **Join Permissions**: After object creation, the script applies domain join permissions automatically by granting full control (GenericAll) to the join group specified in the CSV (using ADSI to update the object’s ACL).  
- **Template Generation**: Run with the `-GenerateTemplate` switch to create a sample CSV template named `ComputerTemplate.csv`.

### Execution Policy

Depending on your setup, you may need to allow script execution:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```
This ensures locally created scripts can run.

---

## Configuration & Customization

- **Suppressing or showing console warnings/errors**: By default, these scripts set `$WarningPreference` and `$ErrorActionPreference` to `SilentlyContinue`. You can change them to `'Continue'` if you prefer to see all warnings/errors in real time.  
- **Transcript Logging**:  
  - If you do **not** specify `-TranscriptPath`, each script creates a timestamped log file in the local directory.  
  - If you **do** specify `-TranscriptPath`, logs go to that file.  
- **CSV Export**:  
  - Use `-Output` if you want the results written to a CSV file. Otherwise, the scripts display results in a formatted table.  
- **Attribute Selection**:  
  - `Get-ADComputers.ps1` uses `-Properties *` but only outputs a subset of fields. You can add/remove attributes based on your needs (`Description`, `Location`, `OperatingSystem`, etc.).  
  - `Get-MacAddresses.ps1` retrieves MAC addresses via a CIM session; you can modify the code to gather more network-related info if desired.  
- **Join Permissions** (in `Create-ADComputers.ps1`):  
  - The script automatically applies domain join permissions by granting full control to the join group specified in the CSV file. This uses an ADSI binding with the LDAP triple-slash syntax to update the object's ACL.

---

## Contributing

1. **Fork this repo**.  
2. Create a new branch for your feature or bugfix:
   ```bash
   git checkout -b feature/add-cool-script
   ```
3. **Commit your changes**:
   ```bash
   git commit -m "Add new script for AD computer creation with join permissions"
   ```
4. **Push to your branch**:
   ```bash
   git push origin feature/add-cool-script
   ```
5. **Open a Pull Request** on GitHub. Describe what your script does and how it improves or fixes an issue.

We appreciate any and all contributions—whether it’s updating documentation, fixing bugs, or adding brand-new functionality.

---

## License

This project is licensed under the **GNU General Public License v3.0**. See the [LICENSE](LICENSE) file for the full text.

---

> **Disclaimer**: Use of these scripts in a production environment should be done cautiously. Always run tests in a **test environment** or **sandbox** before deploying to production to avoid unintended changes to your AD environment.