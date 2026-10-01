# SQL Server Discovery Toolkit

PowerShell scripts for discovering Microsoft SQL Server instances across a Windows estate without requiring a SQL Server login.

The toolkit provides two complementary discovery methods:

- **Active Directory discovery** finds instances that publish `MSSQLSvc` service principal names.
- **Subnet discovery** finds SQL Server instances listening on the network, including instances not registered in Active Directory.

Both scripts enrich discovered instances by reading the Windows registry remotely and export the results to CSV. They recognize SQL Server releases from legacy versions through SQL Server 2025 and report unknown future major versions rather than silently discarding them.

## Safety, support, and warranty

> [!CAUTION]
> These scripts are provided **"AS IS", without warranty of any kind**, express or implied. You run them entirely **at your own risk**. They are community-provided, are **not supported by Microsoft**, and do not constitute official Microsoft guidance.
>
> **Always test the scripts in a non-production/test environment first.** Confirm their behavior, performance, network impact, permissions, and output before considering use in production.
>
> Only query hosts and scan networks for which you have explicit written authorization. Unauthorized network scanning may breach organizational policy or applicable law and may trigger security monitoring.

## Contents

| File | Purpose |
|---|---|
| [`SQL-Discovery-DomainController.ps1`](./SQL-Discovery-DomainController.ps1) | Discovers SQL Server instances through Active Directory SPNs |
| [`SQL-Discovery-MemberServer.ps1`](./SQL-Discovery-MemberServer.ps1) | Scans an IPv4 subnet for SQL Server instances |
| [`SQL-Discovery-Toolkit.md`](./SQL-Discovery-Toolkit.md) | Full design, usage, reference, and troubleshooting guide |
| [`SQL-Discovery-DomainController.sample.csv`](./SQL-Discovery-DomainController.sample.csv) | Sanitized example Active Directory discovery output |
| [`SQL-Discovery-MemberServer.sample.csv`](./SQL-Discovery-MemberServer.sample.csv) | Sanitized example subnet discovery output |

## Choosing a script

| Question | Recommended script |
|---|---|
| Which SQL Server instances are registered in Active Directory? | `SQL-Discovery-DomainController.ps1` |
| What SQL Server instances respond within a known subnet? | `SQL-Discovery-MemberServer.ps1` |
| Which instances may be unregistered or shadow IT? | Run both and compare their output |
| Do I need a SQL Server login? | No; registry enrichment uses Windows administrator access |

The scripts are complementary. Active Directory discovery is fast but can miss instances without an SPN. Subnet discovery can find unregistered instances but is limited to the authorized networks and ports you scan.

## Requirements

### Common requirements

- Windows PowerShell 5.1 or PowerShell 7.
- Local administrator rights on each target host for full version, edition, patch, and registry details.
- WMI/RPC connectivity to target hosts (TCP 135 and the configured RPC dynamic port range).
- DCOM connectivity for legacy Windows hosts that predate WinRM.
- No SQL Server login is required.

Hosts that are found but cannot be enriched are retained in the CSV with an explanatory `Status` value.

### Active Directory discovery

- Run from a domain controller or a domain-joined Windows host.
- Install the RSAT Active Directory PowerShell module.
- Use an account with permission to read Active Directory (normally available to authenticated domain users).

To install the module on supported Windows Server versions:

```powershell
Install-WindowsFeature RSAT-AD-PowerShell
```

To install it on supported Windows client versions:

```powershell
Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
```

### Subnet discovery

- Outbound UDP 1434 for SQL Browser discovery.
- Outbound TCP access to each SQL port being probed (TCP 1433 by default).
- An explicitly authorized IPv4 CIDR range.

## Download and prepare

Either clone the repository:

```powershell
git clone https://github.com/R1WBM/SQL-Version-Discovery.git
Set-Location .\SQL-Version-Discovery
```

Or download and extract the repository ZIP from GitHub. If Windows marks downloaded scripts as blocked, review them first and then unblock only these scripts:

```powershell
Get-ChildItem -Path . -Filter 'SQL-Discovery-*.ps1' | Unblock-File
```

If your organization permits local scripts but the current PowerShell process prevents them from running, consult your administrator rather than weakening a managed execution policy.

## Active Directory discovery

### Basic use

Run with your current Windows identity:

```powershell
.\SQL-Discovery-DomainController.ps1
```

The default report is written to:

```text
.\SQL-Discovery-DomainController.csv
```

### Alternate credentials

Prompt securely for an account that is a local administrator on the target hosts:

```powershell
.\SQL-Discovery-DomainController.ps1 -AuthMode Prompt
```

Or supply an existing `PSCredential` object:

```powershell
$credential = Get-Credential
.\SQL-Discovery-DomainController.ps1 -Credential $credential
```

Do not place plaintext passwords in scripts, shell history, or configuration files.

### Custom output and throttling

```powershell
.\SQL-Discovery-DomainController.ps1 `
    -OutputCsv 'C:\Reports\SQL-Discovery-DomainController.csv' `
    -ThrottleLimit 24
```

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-AuthMode` | `CurrentUser` | Uses the current identity or securely prompts when set to `Prompt` |
| `-Credential` | None | Uses a supplied `PSCredential`; takes precedence over `-AuthMode` |
| `-OutputCsv` | `.\SQL-Discovery-DomainController.csv` | Destination CSV path |
| `-ThrottleLimit` | `16` | Maximum parallel enrichment operations; accepted range is 1-64 |

## Subnet discovery

### Basic use

Specify an authorized IPv4 subnet in CIDR notation:

```powershell
.\SQL-Discovery-MemberServer.ps1 -TargetSubnet '10.20.1.0/24'
```

The default report is written to:

```text
.\SQL-Discovery-MemberServer.csv
```

### Alternate credentials

```powershell
.\SQL-Discovery-MemberServer.ps1 `
    -TargetSubnet '192.168.50.0/24' `
    -AuthMode Prompt
```

Or:

```powershell
$credential = Get-Credential
.\SQL-Discovery-MemberServer.ps1 `
    -TargetSubnet '192.168.50.0/24' `
    -Credential $credential
```

### Multiple SQL ports

Use `-SqlPorts` when authorized SQL Server instances listen on known custom ports:

```powershell
.\SQL-Discovery-MemberServer.ps1 `
    -TargetSubnet '10.20.1.0/24' `
    -SqlPorts 1433,14330,51433
```

### Custom output, throttling, and timeout

```powershell
.\SQL-Discovery-MemberServer.ps1 `
    -TargetSubnet '10.20.0.0/22' `
    -OutputCsv 'C:\Reports\SQL-Discovery-MemberServer.csv' `
    -ThrottleLimit 64 `
    -TimeoutMs 1000
```

### Larger authorized ranges

`-MaxHosts` is a safety brake that prevents a mistyped CIDR prefix from scanning an unexpectedly large network. The default limit is 4,094 usable hosts. Raise it only after verifying the subnet and obtaining authorization:

```powershell
.\SQL-Discovery-MemberServer.ps1 `
    -TargetSubnet '10.20.0.0/20' `
    -MaxHosts 4094
```

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-TargetSubnet` | `10.20.1.0/24` | Authorized IPv4 CIDR range; supported prefixes are `/8` through `/30` |
| `-AuthMode` | `CurrentUser` | Uses the current identity or securely prompts when set to `Prompt` |
| `-Credential` | None | Uses a supplied `PSCredential`; takes precedence over `-AuthMode` |
| `-OutputCsv` | `.\SQL-Discovery-MemberServer.csv` | Destination CSV path |
| `-SqlPorts` | `1433` | One or more TCP ports to probe |
| `-ThrottleLimit` | `32` | Maximum parallel host scans; accepted range is 1-128 |
| `-TimeoutMs` | `700` | Network probe timeout in milliseconds; accepted range is 100-10,000 |
| `-MaxHosts` | `4094` | Maximum number of usable addresses accepted as a scan target |

## Understanding the output

Output columns include:

- Server and instance names
- FQDN and IP address
- SQL Server release, version, patch level, and edition
- Registry instance key and collation
- Discovery method and detection source
- Scan status and timestamp

Review the sample reports before running:

- [`SQL-Discovery-DomainController.sample.csv`](./SQL-Discovery-DomainController.sample.csv)
- [`SQL-Discovery-MemberServer.sample.csv`](./SQL-Discovery-MemberServer.sample.csv)

To compare results by server and instance:

```powershell
$ad = Import-Csv .\SQL-Discovery-DomainController.csv
$network = Import-Csv .\SQL-Discovery-MemberServer.csv

Compare-Object `
    -ReferenceObject ($ad | ForEach-Object { "$($_.FQDN)\$($_.InstanceName)" }) `
    -DifferenceObject ($network | ForEach-Object { "$($_.FQDN)\$($_.InstanceName)" })
```

A result present only in the subnet report may indicate an instance without a registered SPN. Investigate findings rather than assuming they are unauthorized systems.

## Troubleshooting

### Active Directory module cannot be loaded

Install the appropriate RSAT feature and confirm it is available:

```powershell
Get-Module -ListAvailable ActiveDirectory
```

### Access denied or detail unavailable

Confirm that the selected identity:

- Is a local administrator on the target host.
- Can reach WMI/RPC on the target.
- Is not blocked by Windows Firewall, network firewalls, or remote administration policy.

The scripts still report discovered hosts when enrichment fails.

### Named instances are missing

Confirm that:

- UDP 1434 is permitted.
- SQL Browser is running where required.
- Any known custom SQL ports are included with `-SqlPorts`.
- The target subnet is correct and explicitly authorized.

### Legacy hosts cannot be read

Old Windows versions may not support WinRM. The scripts retry with DCOM, but remote WMI/DCOM must still be enabled and permitted through intervening firewalls.

### The subnet is rejected

Check the CIDR notation and resulting host count. Prefixes `/31` and `/32` have no usable host range for this scanner. A large range may exceed `-MaxHosts`; verify it before deliberately raising the limit.

For deeper implementation details and additional troubleshooting, see the [full toolkit guide](./SQL-Discovery-Toolkit.md).

## Protect discovery reports

Generated CSV files can expose server names, domain names, private IP addresses, SQL versions, editions, and patch levels. Treat them as sensitive infrastructure inventories:

- Store reports only in approved locations.
- Apply least-privilege access.
- Do not publish unreviewed reports in issues, pull requests, chat, or this repository.
- Sanitize data before sharing it externally.

The default generated CSV filenames are excluded by [`.gitignore`](./.gitignore) to reduce the chance of committing live environment data. The `.sample.csv` files contain example lab data only.

## License

The source and documentation in this repository are licensed under the [GNU General Public License v3.0 only](./LICENSE) (`GPL-3.0-only`). The license includes explicit warranty and liability disclaimers; the operational safety warnings above still apply.

Earlier revisions released under the MIT License remain available under the terms that accompanied those revisions.
