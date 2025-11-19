# Create AD Computer from CSV (`Create-ADComputer.ps1`)

This script creates Active Directory computer objects from a CSV file and immediately applies domain join permissions and group memberships based on the CSV data. It is meant to standardize how new workstations or servers are created in AD.

---

## Features

- Reads computer definitions from a CSV file.
- Builds computer names based on fields such as site ID and service tag.
- Creates the computer object in the specified OU.
- Adds the computer to one or more AD groups for permissions or targeting.
- Grants a join group full control on the computer object so it can join the domain.
- Writes a transcript log for auditing.
- Can generate a CSV template for you.

---

## Requirements

- PowerShell 5.1 or later.
- Active Directory module:

  ```powershell
  Import-Module ActiveDirectory
  ```

* Permissions to create computer objects in the target OUs and to edit their ACLs.

---

## Parameters

Common parameters (exact names may differ slightly from this example, update to match your script):

* `-Path`
  Path to the input CSV file. If omitted, the script can prompt with a file selection dialog.

* `-GenerateTemplate`
  When present, the script generates a sample CSV file (for example `ComputerTemplate.csv`) with the required columns and exits.

* `-TranscriptPath`
  Optional path for the transcript log. If not supplied, a file such as `Create-ADComputer-YYYYMMDD_HHMMSS.log` is created in the current directory.

---

## CSV format

Adjust this section to match your actual columns. A common layout might look like:

```text
SiteId,ServiceTag,Description,OU,JoinGroup,Groups
ABR1,ABC1234,"Classroom PC 01","OU=Workstations,OU=Education,DC=example,DC=com","Join-Workstations","G-Workstations,G-Physical-Lab"
```

Typical columns:

* `SiteId` - Site or location code.
* `ServiceTag` - Hardware identifier used as part of the name.
* `Description` - AD description field for the computer.
* `OU` - Distinguished name of the OU where the computer object should be created.
* `JoinGroup` - Group that will be granted full control on the computer object for domain join.
* `Groups` - One or more security groups (comma separated) that the computer should be added to.

---

## Usage

### Generate a CSV template

```powershell
.\Create-ADComputer.ps1 -GenerateTemplate
```

This creates a sample CSV you can fill out with your own data.

### Create computers from a CSV file

```powershell
.\Create-ADComputer.ps1 -Path "C:\Temp\computers.csv"
```

The script will:

1. Read each row from the CSV.
2. Build a computer name from the site and service tag (or according to your naming logic).
3. Create the computer object in the specified OU.
4. Add the computer to any groups listed in the CSV.
5. Grant the join group full control on the object to allow domain joins.
6. Write a transcript log.

### Specify a custom transcript path

```powershell
.\Create-ADComputer.ps1 -Path "C:\Temp\computers.csv" -TranscriptPath "C:\Logs\Create-ADComputer.log"
```

---

## Notes

* Consider testing against a lab OU before you run this in production.
* If you change your naming convention or group layout later, you can update the CSV template and reuse the script.