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
| [`SQL-Database-Storage-Inventory.ps1`](./SQL-Database-Storage-Inventory.ps1) | Counts databases and measures actual data and transaction-log storage used |
| [`SQL-Discovery-Toolkit.md`](./SQL-Discovery-Toolkit.md) | Full design, usage, reference, and troubleshooting guide |
| [`SQL-Discovery-DomainController.sample.csv`](./SQL-Discovery-DomainController.sample.csv) | Sanitized example Active Directory discovery output |
| [`SQL-Discovery-MemberServer.sample.csv`](./SQL-Discovery-MemberServer.sample.csv) | Sanitized example subnet discovery output |

## Choosing a script

| Question | Recommended script |
|---|---|
| Which SQL Server instances are registered in Active Directory? | `SQL-Discovery-DomainController.ps1` |
| What SQL Server instances respond within a known subnet? | `SQL-Discovery-MemberServer.ps1` |
| Which instances may be unregistered or shadow IT? | Run both and compare their output |
| Do discovery scans need a SQL Server login? | No; discovery enrichment uses Windows remote-registry access |
| Does database storage inventory need a SQL Server login? | Yes; the Windows or SQL identity must exist as a login on every queried SQL instance |

The scripts are complementary. Active Directory discovery is fast but can miss instances without an SPN. Subnet discovery can find unregistered instances but is limited to the authorized networks and ports you scan.

## Database storage inventory

### How it connects

The storage inventory uses .NET `System.Data.SqlClient.SqlConnection` to make a
normal SQL Server connection from the member server. It does not use remote
PowerShell, WMI, or the remote registry for database measurements.

Connection targets are derived from each discovery CSV row:

- A detected TCP port becomes `tcp:FQDN,port`.
- A named instance becomes `FQDN\InstanceName`.
- A default instance without a detected port uses its FQDN or server name.

The member server must be able to resolve the target name and reach the SQL
Server listening port through intervening firewalls. Named instances without a
known port may also require SQL Server Browser/UDP 1434.

Authentication behavior is:

1. **Windows Integrated Authentication is the default.** If
   `-Credential` is omitted and `-AuthMode CurrentUser` is active,
   `Integrated Security=True` is used. SQL Server sees the Windows identity
   running PowerShell. Confirm it before running:

   ```powershell
   whoami
   ```

2. **SQL authentication is used when `-AuthMode Prompt` or `-Credential` is
   supplied.** Prompt mode securely requests a SQL login and password.
   `-Credential` accepts an existing `PSCredential` and takes precedence over
   `-AuthMode`. The earlier `-SqlCredential` parameter name remains available
   as an alias. These options do not impersonate another Windows account.

3. **The script does not silently switch authentication methods after a failed
   login.** It uses Windows authentication for the entire run when the default
   `-AuthMode CurrentUser` is active and no credential was supplied. It uses
   SQL authentication for the entire run when prompt mode or an explicit
   credential supplies a SQL login.

To use a different Windows identity, launch PowerShell as that identity and run
the script with the default `-AuthMode CurrentUser`. A SQL connection cannot
use a `PSCredential` as an alternate Windows identity merely by placing it in a
connection string.

Encryption is not explicitly requested by default. Use `-Encrypt` to request an
encrypted connection. Use `-TrustServerCertificate` only when permitted by your
security policy because it bypasses certificate-chain validation.

### Permissions required in SQL Server

Local administrator rights on the Windows host do not automatically grant SQL
Server access. The selected Windows or SQL identity must be configured as a
login separately on every SQL Server instance being inventoried.

The account needs enough permission to:

- Connect to the instance and the databases being measured.
- Enumerate databases.
- read database file metadata and `FILEPROPERTY(..., 'SpaceUsed')`.
- Run `DBCC SQLPERF(LOGSPACE)`.

Use a dedicated least-privilege account where possible. A DBA should review and
run the following examples on each target instance. Do not grant `sysadmin`
merely to make the inventory work.

#### Grant access to a Windows account

Replace `CONTOSO\SqlInventory` with the domain account that will run the
PowerShell script:

```sql
USE [master];
GO
CREATE LOGIN [CONTOSO\SqlInventory] FROM WINDOWS;
GO
GRANT VIEW ANY DATABASE TO [CONTOSO\SqlInventory];
GRANT CONNECT ANY DATABASE TO [CONTOSO\SqlInventory];
GO
```

`CONNECT ANY DATABASE` is available in SQL Server 2014 and newer. On older
versions, or where policy does not permit that server-level permission, create a
database user and grant access separately in every database that should be
measured:

```sql
USE [YourDatabase];
GO
CREATE USER [CONTOSO\SqlInventory] FOR LOGIN [CONTOSO\SqlInventory];
GRANT CONNECT TO [CONTOSO\SqlInventory];
GRANT VIEW DATABASE STATE TO [CONTOSO\SqlInventory];
GO
```

Repeat that database-level block for each included user and system database.
If the user already exists, omit the relevant `CREATE USER` statement.

`DBCC SQLPERF(LOGSPACE)` also needs a server-state permission. Use the statement
appropriate for the SQL Server version:

```sql
-- SQL Server 2019 and earlier
GRANT VIEW SERVER STATE TO [CONTOSO\SqlInventory];
GO

-- SQL Server 2022 and later
GRANT VIEW SERVER PERFORMANCE STATE TO [CONTOSO\SqlInventory];
GO
```

Run only the version-appropriate grant, unless your DBA determines both are
required for the versions and security policy in the estate.

#### Create and grant access to a SQL login

SQL authentication requires SQL Server mixed mode to be enabled. A DBA can
create a dedicated SQL login using the organization's password-management
process:

```sql
USE [master];
GO
CREATE LOGIN [SqlStorageInventory]
WITH PASSWORD = 'ReplaceWithAStrongManagedPassword',
     CHECK_POLICY = ON,
     CHECK_EXPIRATION = ON;
GO
GRANT VIEW ANY DATABASE TO [SqlStorageInventory];
GRANT CONNECT ANY DATABASE TO [SqlStorageInventory];
GO
```

Grant the same version-appropriate server-state permission shown above, replacing
`[CONTOSO\SqlInventory]` with `[SqlStorageInventory]`. Where database-level users
are required, create one in every database being measured:

```sql
USE [YourDatabase];
GO
CREATE USER [SqlStorageInventory] FOR LOGIN [SqlStorageInventory];
GRANT CONNECT TO [SqlStorageInventory];
GRANT VIEW DATABASE STATE TO [SqlStorageInventory];
GO
```

Never place the SQL password in this repository, a script, a CSV, command
history, or the diagnostic file.

### Run the inventory

After producing one or both discovery CSV files, copy the inventory script and
CSV files to an authorised member server. Run it with the current Windows
identity:

```powershell
.\SQL-Database-Storage-Inventory.ps1 `
    -InputCsv '.\SQL-Discovery-DomainController.csv',
              '.\SQL-Discovery-MemberServer.csv'
```

The script deduplicates instances present in both files. Actual used storage
includes occupied data pages and used transaction-log space. System databases
are included unless `-ExcludeSystemDatabases` is supplied:

```powershell
.\SQL-Database-Storage-Inventory.ps1 -ExcludeSystemDatabases
```

For SQL authentication, obtain the password through a secure interactive prompt
using either prompt mode:

```powershell
.\SQL-Database-Storage-Inventory.ps1 -AuthMode Prompt
```

Or create a credential object and supply it directly. An explicit credential
overrides `-AuthMode`:

```powershell
$credential = Get-Credential
.\SQL-Database-Storage-Inventory.ps1 -Credential $credential
```

To request encryption:

```powershell
.\SQL-Database-Storage-Inventory.ps1 -Encrypt
```

If an authorised test environment uses a certificate that the member server
does not trust, and policy permits bypassing certificate validation:

```powershell
.\SQL-Database-Storage-Inventory.ps1 -Encrypt -TrustServerCertificate
```

### Inventory parameters

| Parameter | Default | Description |
|---|---|---|
| `-InputCsv` | Both standard discovery CSV names | One or more discovery CSV files |
| `-OutputCsv` | `.\SQL-Database-Storage-Inventory.csv` | Results CSV containing every instance and its status |
| `-DiagnosticLog` | `.\SQL-Database-Storage-Inventory-Diagnostics.txt` | Append-only error details; created or modified only when a warning or error occurs |
| `-AuthMode` | `CurrentUser` | Uses the current Windows identity, or securely prompts for a SQL login when set to `Prompt` |
| `-Credential` | None | Uses a supplied SQL-login `PSCredential` and takes precedence over `-AuthMode`; `-SqlCredential` is retained as an alias |
| `-ExcludeSystemDatabases` | Off | Excludes `master`, `model`, `msdb`, and `tempdb` |
| `-ConnectionTimeoutSeconds` | `10` | SQL connection timeout per instance |
| `-CommandTimeoutSeconds` | `30` | SQL command timeout per query |
| `-Encrypt` | Off | Requests SQL transport encryption |
| `-TrustServerCertificate` | Off | Bypasses certificate-chain validation when encryption is used |

### Results and diagnostics

The default report is written to
`.\SQL-Database-Storage-Inventory.csv`. Instances or databases that cannot be
queried remain in the report with `FAILED` or `PARTIAL` status.

Detailed diagnostics are appended to
`.\SQL-Database-Storage-Inventory-Diagnostics.txt` by default. Console warnings
include diagnostic IDs that identify the matching text-file entries. Error
details are not written to the inventory CSV. The text file is created or
appended only when at least one warning or error occurs; a fully successful run
does not create or modify it. Use
`-DiagnosticLog 'C:\Reports\SQL-Storage-Diagnostics.txt'` to select a different
location. The text file includes timestamps, exception and inner-exception
details, PowerShell and .NET stack traces, SQL error numbers, operation context,
the executing identity, and the input and output paths. Processing continues
after individual server or database failures, and every discovered instance
remains in the CSV with its `OK`, `PARTIAL`, or `FAILED` status.

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
