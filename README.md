# Active Directory and Domain Management PowerShell Scripts

A small collection of PowerShell scripts to help with day to day Active Directory (AD) and domain joined computer management. Each script lives in its own folder with a dedicated README that covers parameters, examples, and usage notes.

This root README gives you a high level overview and links to each script.

---

## Scripts

| Script | Short description | Typical use cases | Folder |
| ------ | ------------------ | ----------------- | ------ |
| `Create-ADComputer.ps1` | Create computer objects from a CSV and apply join permissions automatically. | New workstation rollouts, pre staging computers in AD, applying standard group memberships and join permissions. | [`Create-ADComputer/`](Create-ADComputer/) |
| `Get-ADComputers.ps1` | Enumerate AD computer objects from an OU and export key attributes. | Inventory of servers or workstations, basic audit of computer objects, feeding other tooling with a list of machines. | [`Get-ADComputers/`](Get-ADComputers/) |
| `Get-MACAddresses.ps1` | Query AD computers in an OU, check which are online, and collect MAC addresses. | Building IP/MAC inventories, preparing DHCP reservations, mapping hardware on the network. | [`Get-MACAddresses/`](Get-MACAddresses/) |

Each script can be used on its own or combined in a workflow. For example, you might:

1. Use `Get-ADComputers.ps1` to export an inventory of servers.
2. Use `Get-MACAddresses.ps1` to fill in MAC addresses for online hosts.
3. Use `Create-ADComputer.ps1` for new builds that need to follow a standard naming convention and group membership.

---

## Requirements

All scripts assume:

- PowerShell 5.1 or later.
- The Active Directory module is available:

  ```powershell
  Import-Module ActiveDirectory
````

* You are running with sufficient permissions to read or modify the AD objects that the script touches.
* Network connectivity and firewall rules permit access to your domain controllers and the target computers (for scripts that query the machines themselves).

---

## General usage notes

* **Execution policy**

  You may need to allow local scripts to run:

  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

* **Transcript logging**

  Each script is designed to log its work to a transcript file. By default a timestamped transcript is created in the current directory if you do not specify a custom path. See each script README for details.

* **CSV export**

  Where applicable, scripts support an `-Output` parameter to export results to CSV. If `-Output` is not provided, results are shown in a formatted table in the console.

* **Error and warning handling**

  Scripts are written to reduce noisy console output by adjusting `$WarningPreference` and `$ErrorActionPreference`. You can adjust those if you prefer more verbose feedback during execution.

---

## Contributing

If you have ideas for additional scripts or improvements:

* Open an issue with a description of what you would like to add or change.
* Or submit a pull request that:

  * Keeps the per script folder structure.
  * Includes a README for any new script.

---

## License

This project is licensed under the GPL 3.0 license. See [`LICENSE`](LICENSE) for details.