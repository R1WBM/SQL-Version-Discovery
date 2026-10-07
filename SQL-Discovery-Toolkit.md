# SQL Server Discovery Toolkit

**Author:** Russell McKee
**LinkedIn:** https://www.linkedin.com/in/russellwbmckee/
**Version:** 2.3
**Date:** 7 October 2026

---

## Change log

| Version | Change |
|---|---|
| **2.3** | Added `SQL-Database-Storage-Inventory.ps1` — a third, optional script that expands discovery output into individual SQL Server instances and reports SQL-visible vCore counts, database counts, and actual data and transaction-log storage used. Unlike the two discovery scripts, it requires a SQL Server login (Windows or SQL) on every queried instance. Adds instance expansion via SQL Browser and WMI/remote registry, per-instance connectivity diagnostics (including a `NOT CONNECTABLE` status for instances with every remote protocol disabled), and a separate append-only diagnostic log. |
| **2.1** | Genuine support for legacy SQL Server (7.0, 2000, 2005) — the SQL 2000 default-instance hive, `Wow6432Node` redirection, `CurrentVersion` fallback where `PatchLevel` doesn't exist, and a DCOM transport fallback for hosts predating WinRM. Version map extended down to SQL 6.5. |
| **2.0** | Replaced SQL-login discovery with remote registry reads (no SQL permissions needed); removed the hard-coded SQL 2016 version filter; added `-AuthMode` and `-Credential`; parallel subnet sweep via runspace pool. |
| **1.x** | Original scripts — SQL-login based, hard-coded to SQL 2016. |

---

## Disclaimer

These scripts are provided **"AS IS"**, without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose and non-infringement.

In no event shall the author be liable for any claim, damages, loss of data, loss of service, or other liability, whether in an action of contract, tort or otherwise, arising from, out of, or in connection with the use of these scripts.

You run these scripts entirely at your own risk. They are **not supported by Microsoft** and do not constitute official Microsoft guidance.

**Always test in a non-production / test environment first.** Satisfy yourself that the behaviour is appropriate before running against production systems, and ensure you have written authorisation to query (and, for the subnet scanner, to scan) the systems in scope.

---

## 1. Overview

This toolkit contains three complementary PowerShell scripts. Two discover Microsoft SQL Server instances across an estate; the third expands that discovery into a per-instance vCore, database, and storage inventory. All three export their results to CSV.

| Script | Runs from | Finds SQL by | Best for |
|---|---|---|---|
| `SQL-Discovery-DomainController.ps1` | A domain controller (or any host with RSAT) | Querying Active Directory for `MSSQLSvc` SPNs | Fast, authoritative sweep of everything registered in AD |
| `SQL-Discovery-MemberServer.ps1` | Any domain member server | Sweeping every IP in a subnet | Finding instances AD doesn't know about |
| `SQL-Database-Storage-Inventory.ps1` | Any domain member server, consuming the CSV output of the two discovery scripts | Connecting directly to each discovered SQL instance | SQL-visible vCore counts, database counts, and actual data/transaction-log storage used |

The two discovery scripts capture the same core fields: **Server Name, Instance, SQL Version, SQL Edition, FQDN and IP Address** — for **any SQL Server release from 6.5 / 7.0 / 2000 through to 2025**, with no hard-coded version filter.

### Why two discovery scripts?

They answer different questions, and the gap between them is the interesting part.

The **domain controller** script asks *"what does Active Directory think we have?"*. It's near-instant (1.3 seconds in the lab) because it's a single LDAP query, but it only sees instances that registered a Service Principal Name. An instance running under `LocalSystem`, or one whose SPN registration failed, is invisible to it.

The **member server** script asks *"what's actually listening on the wire?"*. It's slower (13 seconds for a /24) but it doesn't care about AD registration — if something answers on a SQL port, it gets found.

**Run both and compare.** Anything the subnet scan finds that the AD scan missed is either an unregistered instance or shadow IT, and both are worth knowing about.

### Why a separate storage inventory script?

Discovery deliberately needs no SQL permissions at all (see section 2). Sizing an estate — vCores, database counts, actual data and transaction-log storage — needs a real SQL Server connection, which is a different permission system entirely. Keeping that a separate, optional third step means the two discovery scripts stay usable by anyone with local administrator rights on the target hosts, while storage inventory remains an explicit, auditable action that a DBA grants access for separately. See section 3 for the exact least-privilege grants.

---

## 2. The key design decision: no SQL login required

Most SQL discovery scripts connect to each instance and run `SELECT SERVERPROPERTY(...)`. That fails the moment you hit an instance where your account has no SQL login — even if you're a Domain Admin. Being an administrator of the *Windows server* grants you nothing inside the *SQL engine*; they are separate permission systems.

That exact failure is what these scripts were written to solve. During development, the original version returned zero results in a working lab with a live SQL Server, failing with:

```
Login failed for user 'SQLTEST\azureadmin'
```

...despite that account being a Domain Admin.

**The fix: read the Windows registry instead of querying SQL.**

SQL Server writes its version, patch level and edition into the registry at install time:

```
HKLM\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL       <- instance list
HKLM\SOFTWARE\Microsoft\Microsoft SQL Server\<InstanceKey>\Setup      <- version, edition, patch
```

Reading those needs **local administrator on the target server** — which a Domain Admin already has — and **no SQL permissions whatsoever**. The SQL service doesn't even need to be running.

This is done via WMI's `StdRegProv` class over a CIM session, which (unlike `OpenRemoteBaseKey`) accepts alternate credentials, so the same code path works whether you're using your current identity or a supplied one.

### Legacy SQL versions (2000, 2005 and earlier)

The registry layout above is the *modern* one. Older releases stored things differently, so the scripts probe a series of candidate paths and take the first that yields a build number:

| Release | Instance list | Version location |
|---|---|---|
| SQL 2005 – 2025 | `Instance Names\SQL` | `Microsoft SQL Server\MSSQL##.INSTANCE\Setup\Version` |
| SQL 2000 (named) | `Instance Names\SQL` | `Microsoft SQL Server\<InstanceName>\MSSQLServer\CurrentVersion` |
| SQL 2000 (default) | **not listed** | `Microsoft\MSSQLServer\MSSQLServer\CurrentVersion` — a different hive entirely |
| SQL 7.0 | not listed | as above |

Three specific accommodations are made for old estates:

1. **The SQL 2000 default instance is probed directly.** It lives under `HKLM\SOFTWARE\Microsoft\MSSQLServer` (no "Microsoft SQL Server" parent) and doesn't reliably appear in `Instance Names\SQL`. If the modern enumeration returns nothing, that hive is checked before giving up.

2. **`Wow6432Node` is checked as well as the native view.** 32-bit SQL on 64-bit Windows was common in the 2000/2005 era, and its keys are redirected.

3. **`PatchLevel` doesn't exist before SQL 2008.** Where it's absent, the script falls back to the running `CurrentVersion`, then to the install `Version`, so the column is never empty for an instance it could read.

There's also a transport problem with genuinely old hosts: **WinRM didn't ship until Windows Server 2008**, so a `New-CimSession` to a Windows 2000/2003 box fails outright. Both scripts catch that and **retry over DCOM**, which those systems do support.

> **Reality check on very old estates.** SQL 2000 and 2005 are long out of support and typically sit on equally old Windows. Remote WMI may be disabled or firewalled on such hosts. Where the scripts can't read a host they still report it with a `Status` explaining why — you get a lead to chase rather than a silent omission.

> **This section applies to the two discovery scripts only.** `SQL-Database-Storage-Inventory.ps1` is the deliberate exception — it connects to SQL Server directly to read database and log space, so it does need a SQL Server login (Windows or SQL) on every instance it measures. Local administrator rights on the Windows host do not substitute for that login. See section 3 for the specific grants required.

---

## 3. Access required

### To run the domain controller script

| Requirement | Why | If missing |
|---|---|---|
| Read access to Active Directory | Query `MSSQLSvc` SPNs | Any authenticated domain user has this by default |
| **RSAT ActiveDirectory PowerShell module** | `Get-ADObject` | Script stops with a clear message telling you to install `RSAT-AD-PowerShell` |
| **Local Administrator on each target SQL host** | Read the registry remotely | That host is still reported, with `Status` explaining why detail is missing |
| WMI / RPC to targets (TCP 135 + dynamic range) | Remote registry transport | Host reported as unreachable |
| *(legacy hosts only)* DCOM enabled on the target | Windows 2000/2003 predates WinRM, so the script falls back to DCOM | Old SQL host found but not readable |

### To run the member server script

| Requirement | Why | If missing |
|---|---|---|
| **UDP 1434 outbound** | SQL Browser (SSRP) discovery | Falls back to TCP port probe |
| **TCP 1433 outbound** (and any custom ports) | Port probe | Instance won't be found at all |
| **Local Administrator on each target SQL host** | Read the registry remotely | Host reported with `(port open, detail unavailable)` |
| WMI / RPC to targets | Remote registry transport | As above |
| *(legacy hosts only)* DCOM enabled on the target | Windows 2000/2003 predates WinRM, so the script falls back to DCOM | Old SQL host found but not readable |
| **Authorisation to scan the subnet** | It is a network scan | Don't run it |

### Does it need to run elevated?

Yes — run PowerShell **as Administrator**. The elevation is needed locally to establish the outbound CIM/WMI sessions.

### A note on least privilege

Domain Admin is *sufficient* but not *required*. The actual requirement is local administrator on the SQL hosts. In a tightly-governed environment, create a dedicated discovery account, add it to the local Administrators group on the SQL servers only, and pass it with `-Credential`. That gives you the inventory without handing out domain-wide rights.

### To run the database storage inventory script

| Requirement | Why | If missing |
|---|---|---|
| Output CSV from one or both discovery scripts | Input — the inventory expands hosts into instances rather than discovering them itself | Script looks for both standard discovery filenames by default; supply `-InputCsv` otherwise |
| Network path to each instance's SQL port | Direct SQL connection via .NET `System.Data.SqlClient.SqlConnection` | Instance reported `NOT CONNECTABLE` or `FAILED`, with detail in the diagnostic log |
| **A SQL Server login on every target instance** | The script runs real T-SQL (`FILEPROPERTY(..., 'SpaceUsed')`, `DBCC SQLPERF(LOGSPACE)`) | Instance reported `FAILED` — local Windows administrator rights do **not** substitute for a SQL login |
| `VIEW ANY DATABASE` + `CONNECT ANY DATABASE` (or per-database `CONNECT` + `VIEW DATABASE STATE`) | Enumerate and read databases | Affected databases are skipped and the instance is reported `PARTIAL` |
| `VIEW SERVER STATE` (SQL 2019 and earlier) or `VIEW SERVER PERFORMANCE STATE` (SQL 2022 and later) | Required by `DBCC SQLPERF(LOGSPACE)` | Transaction-log usage cannot be measured |
| *(optional)* Local Administrator / WMI access to the host | Expands hosts into instances via SQL Browser and the remote registry | Skip entirely with `-InstanceDiscoveryAuthMode BrowserOnly`, which falls back to SQL Browser and the instances already listed in the input CSVs |

**This is the one script in the toolkit that needs a SQL Server login.** Never grant `sysadmin` merely to make it work — a DBA should create and grant a dedicated least-privilege account instead.

Grant access to a Windows account (replace `CONTOSO\SqlInventory` with the account that will run the script):

```sql
USE [master];
GO
CREATE LOGIN [CONTOSO\SqlInventory] FROM WINDOWS;
GO
GRANT VIEW ANY DATABASE TO [CONTOSO\SqlInventory];
GRANT CONNECT ANY DATABASE TO [CONTOSO\SqlInventory];
GO

-- SQL Server 2019 and earlier
GRANT VIEW SERVER STATE TO [CONTOSO\SqlInventory];
GO

-- SQL Server 2022 and later
GRANT VIEW SERVER PERFORMANCE STATE TO [CONTOSO\SqlInventory];
GO
```

`CONNECT ANY DATABASE` needs SQL Server 2014 or newer. On older versions, or where policy does not permit that server-level permission, create a database user and grant `CONNECT` plus `VIEW DATABASE STATE` in every database to be measured instead. For SQL authentication, a DBA creates a dedicated SQL login (`CREATE LOGIN ... WITH PASSWORD = ...`) and grants it the same permissions shown above. Never place a SQL password in this repository, a script, a CSV, command history, or the diagnostic log.

---

## 4. Usage

### Domain controller script

```powershell
# Simplest form - current identity, CSV in the current folder
.\SQL-Discovery-DomainController.ps1

# Prompt for alternate credentials
.\SQL-Discovery-DomainController.ps1 -AuthMode Prompt

# Non-interactive (scheduled task / automation)
$cred = Get-Credential
.\SQL-Discovery-DomainController.ps1 -Credential $cred

# Custom output location
.\SQL-Discovery-DomainController.ps1 -OutputCsv '\\fileserver\reports\sql-estate.csv'
```

### Member server script

```powershell
# Scan a /24
.\SQL-Discovery-MemberServer.ps1 -TargetSubnet '10.20.1.0/24'

# Prompt for credentials
.\SQL-Discovery-MemberServer.ps1 -TargetSubnet '192.168.50.0/24' -AuthMode Prompt

# Non-interactive
$cred = Get-Credential
.\SQL-Discovery-MemberServer.ps1 -TargetSubnet '10.10.0.0/22' -Credential $cred

# Larger subnet, more parallelism, extra SQL ports
.\SQL-Discovery-MemberServer.ps1 -TargetSubnet '10.0.0.0/22' `
    -ThrottleLimit 64 -SqlPorts 1433,1434,49172 -MaxHosts 2000
```

### Scanning multiple subnets

```powershell
$subnets = '10.20.1.0/24', '10.20.2.0/24', '192.168.10.0/24'
$cred = Get-Credential

$all = foreach ($s in $subnets) {
    .\SQL-Discovery-MemberServer.ps1 -TargetSubnet $s -Credential $cred `
        -OutputCsv ".\scan-$($s -replace '[./]','-').csv"
    Import-Csv ".\scan-$($s -replace '[./]','-').csv"
}
$all | Export-Csv '.\SQL-Estate-Combined.csv' -NoTypeInformation
```

### Database storage inventory script

```powershell
# Simplest form - reads both standard discovery CSVs, current Windows identity
.\SQL-Database-Storage-Inventory.ps1

# Explicit input CSVs
.\SQL-Database-Storage-Inventory.ps1 `
    -InputCsv '.\SQL-Discovery-DomainController.csv', '.\SQL-Discovery-MemberServer.csv'

# Exclude system databases from the storage totals
.\SQL-Database-Storage-Inventory.ps1 -ExcludeSystemDatabases

# SQL authentication via a secure prompt
.\SQL-Database-Storage-Inventory.ps1 -AuthMode Prompt

# SQL authentication via an explicit credential (non-interactive)
$sqlCred = Get-Credential
.\SQL-Database-Storage-Inventory.ps1 -Credential $sqlCred

# Separate Windows discovery credential and SQL login credential
$windowsCred = Get-Credential -Message 'Windows discovery account'
$sqlCred     = Get-Credential -Message 'SQL login'
.\SQL-Database-Storage-Inventory.ps1 -DiscoveryCredential $windowsCred -Credential $sqlCred

# Skip authenticated WMI/registry discovery entirely
.\SQL-Database-Storage-Inventory.ps1 -InstanceDiscoveryAuthMode BrowserOnly

# Request encryption (add -TrustServerCertificate only if policy permits it)
.\SQL-Database-Storage-Inventory.ps1 -Encrypt
```

---

## 5. Variables to change for a different environment

### Domain controller script

| Parameter | Default | Change it when |
|---|---|---|
| `$AuthMode` | `CurrentUser` | Your logged-on account isn't a local admin on the SQL hosts. Use `Prompt`. |
| `$Credential` | *(none)* | Running unattended — a prompt would block a scheduled task. |
| `$OutputCsv` | `.\SQL-Discovery-DomainController.csv` | You want the report on a share or in a dated folder. |

There is **no domain name variable** — the script discovers the current domain automatically, so it's portable as-is.

**You will also want to extend `$VersionMap`** when a new SQL release ships. It's a plain hashtable near the top:

```powershell
$VersionMap = @{
    '6'  = 'SQL Server 6.0 / 6.5'
    '7'  = 'SQL Server 7.0'
    '8'  = 'SQL Server 2000'
    '9'  = 'SQL Server 2005'
    '10' = 'SQL Server 2008 / 2008 R2'
    '11' = 'SQL Server 2012'
    '12' = 'SQL Server 2014'
    '13' = 'SQL Server 2016'
    '14' = 'SQL Server 2017'
    '15' = 'SQL Server 2019'
    '16' = 'SQL Server 2022'
    '17' = 'SQL Server 2025'
    '18' = 'SQL Server vNext'    # <-- add future releases here
}
```

An unrecognised major build is **not** discarded — it's reported as `Unrecognised (major build 18)`, so the script never silently loses an instance just because it predates your version map. (This matters: the original script hard-coded a `13.*` filter for SQL 2016, which silently discarded every SQL 2022 instance in the estate.)

**Adding an unusual registry layout.** If you meet an estate where an instance is listed but no version resolves, the candidate paths live in `Resolve-SqlInstanceDetail` as a simple ordered list:

```powershell
$roots = [System.Collections.Generic.List[string]]::new()
if ($InstanceKey)  { $roots.Add("SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceKey") }
if ($InstanceName -and $InstanceName -ne $InstanceKey) {
    $roots.Add("SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceName")
}
$roots.Add('SOFTWARE\Microsoft\MSSQLServer')       # SQL 2000 default instance
# $roots.Add('SOFTWARE\Vendor\CustomPath')         # <-- add a new layout here
```

Each entry is automatically also probed under `Wow6432Node`, so you only add the native path. The first path that returns a build number wins.

### Member server script

| Parameter | Default | Change it when |
|---|---|---|
| **`$TargetSubnet`** | `10.20.1.0/24` | **Always** — set this to the subnet you want to scan. |
| `$AuthMode` | `CurrentUser` | Your account isn't a local admin on the targets. |
| `$Credential` | *(none)* | Running unattended. |
| `$SqlPorts` | `@(1433)` | You use non-default ports, e.g. `@(1433,1434,49172)`. |
| `$ThrottleLimit` | `32` | Larger subnet (raise to 64) or a congested link (lower to 8-16). |
| `$TimeoutMs` | `700` | High-latency links — raise to 1500-2000. Lower for fast LANs. |
| `$MaxHosts` | `4094` | Safety brake. Scanning anything larger than a `/20` needs a deliberate raise. |
| `$OutputCsv` | `.\SQL-Discovery-MemberServer.csv` | Reporting to a share. |

#### CIDR examples

| Subnet | Usable hosts | Approx. scan time |
|---|---|---|
| `10.20.1.0/24` | 254 | ~13 seconds |
| `10.20.0.0/22` | 1,022 | ~50 seconds |
| `192.168.1.0/25` | 126 | ~7 seconds |
| `172.16.0.0/20` | 4,094 | ~3 minutes |

`/31` and `/32` are rejected — they have no conventional usable host range.

### Database storage inventory script

| Parameter | Default | Change it when |
|---|---|---|
| `-InputCsv` | Both standard discovery CSV names | Your discovery CSVs use custom names or locations |
| `-OutputCsv` | `.\SQL-Database-Storage-Inventory.csv` | Reporting to a share or dated folder |
| `-DiagnosticLog` | `.\SQL-Database-Storage-Inventory-Diagnostics.txt` | You want detailed error/warning text kept elsewhere; it's only created or appended when a warning or error occurs |
| `-AuthMode` | `CurrentUser` | The current Windows identity has no SQL login. Use `Prompt`, or supply `-Credential` for unattended runs |
| `-Credential` | *(none)* | Running unattended with a SQL login; takes precedence over `-AuthMode` (`-SqlCredential` is an alias) |
| `-InstanceDiscoveryAuthMode` | `CurrentUser` | The identity running the script can't read the remote registry on targets. Use `Prompt`, or `BrowserOnly` to skip WMI/registry entirely |
| `-DiscoveryCredential` | *(none)* | Non-interactive Windows discovery credential, separate from the SQL login (`-WindowsCredential` / `-InstanceDiscoveryCredential` are aliases) |
| `-InstanceDiscoveryTimeoutMs` | `1000` | SQL Browser is slow to respond on a congested or high-latency network |
| `-ExcludeSystemDatabases` | Off | You only want user-database storage, excluding `master`, `model`, `msdb`, `tempdb` |
| `-ConnectionTimeoutSeconds` | `10` | High-latency links need longer to establish a SQL connection |
| `-CommandTimeoutSeconds` | `30` | Very large databases need longer for `FILEPROPERTY`/`DBCC SQLPERF` to return |
| `-Encrypt` | Off | Security policy requires encrypted SQL connections |
| `-TrustServerCertificate` | Off | An authorised test environment uses a certificate the member server doesn't trust, and policy permits bypassing chain validation |

Unlike the two discovery scripts, there is no subnet or domain variable to set — targets come entirely from the discovery CSVs supplied via `-InputCsv`.

---

## 6. Output

Both scripts write a UTF-8 CSV and print the same table to the console.

### Fields

| Column | Description |
|---|---|
| `ServerName` | Short NetBIOS-style host name |
| `InstanceName` | `MSSQLSERVER` for default, otherwise the named instance |
| `FQDN` | Fully qualified domain name (via DNS) |
| `IPAddress` | IPv4 address(es) |
| `SQLRelease` | Friendly release name, e.g. `SQL Server 2022` |
| `SQLVersion` | RTM build, e.g. `16.0.1000.6` |
| `PatchLevel` | Current patched build, e.g. `16.0.4255.1` |
| `SQLEdition` | e.g. `Developer Edition`, `Enterprise Edition` |
| `EditionType` | Edition sub-type |
| `Collation` | Server default collation |
| `InstanceKey` | Registry hive, e.g. `MSSQL16.MSSQLSERVER` |
| `ServicePrincipalName` | *(DC script only)* the SPNs that led to this host |
| `DetectedBy` | *(subnet script only)* `SQL Browser (UDP 1434)` or `TCP port 1433` |
| `DiscoveredVia` | Which method produced the row — your audit trail |
| `Status` | `OK`, or a plain-English reason detail is missing |
| `ScanTime` | ISO-8601 timestamp |

> **`SQLVersion` vs `PatchLevel`** — `SQLVersion` is the build the instance was *installed* from; `PatchLevel` is what it is running *now*. For patch compliance reporting, use **`PatchLevel`**. In the sample below the instance was installed from RTM (`16.0.1000.6`) and has since been patched to `16.0.4255.1`.

### Example output — domain controller script

```
"ServerName","InstanceName","FQDN","IPAddress","SQLRelease","SQLVersion","PatchLevel","SQLEdition","EditionType","Collation","InstanceKey","ServicePrincipalName","DiscoveredVia","Status","ScanTime"
"SQLTEST-SQL01","MSSQLSERVER","sqltest-sql01.sqltest.local","10.20.1.6","SQL Server 2022","16.0.1000.6","16.0.4255.1","Developer Edition","Developer Edition","SQL_Latin1_General_CP1_CI_AS","MSSQL16.MSSQLSERVER","MSSQLSvc/sqltest-sql01.sqltest.local:1433; MSSQLSvc/sqltest-sql01.sqltest.local","Active Directory SPN + Remote Registry","OK","2026-09-25T20:24:43"
```

### Example output — member server script

```
"ServerName","InstanceName","FQDN","IPAddress","SQLRelease","SQLVersion","PatchLevel","SQLEdition","EditionType","Collation","InstanceKey","DetectedBy","DiscoveredVia","Status","ScanTime"
"SQLTEST-SQL01","MSSQLSERVER","SQLTEST-SQL01.sqltest.local","10.20.1.6","SQL Server 2022","16.0.1000.6","16.0.4255.1","Developer Edition","Developer Edition","SQL_Latin1_General_CP1_CI_AS","MSSQL16.MSSQLSERVER","TCP port 1433","Subnet scan + Remote Registry","OK","2026-09-25T19:55:29"
```

### Console output

```
================================================================
 SQL Server Discovery - Member Server Subnet Scan (v2.1)
 Author: Russell McKee
================================================================
 Target subnet : 10.20.1.0/24
 Auth mode     : Explicit credential (SQLTEST\azureadmin)
 Running as    : NT AUTHORITY\SYSTEM
 SQL ports     : 1433
 Output CSV    : F:\Data\SQL-Discovery-MemberServer.csv

Sweeping 254 usable address(es)...
Found 1 responding SQL endpoint(s).

ServerName    InstanceName SQLRelease      SQLVersion  PatchLevel  SQLEdition
----------    ------------ ----------      ----------  ----------  ----------
SQLTEST-SQL01 MSSQLSERVER  SQL Server 2022 16.0.1000.6 16.0.4255.1 Developer Edition

CSV exported to  : F:\Data\SQL-Discovery-MemberServer.csv
Addresses swept  : 254
Endpoints found  : 1
Instances read   : 1
Elapsed          : 13.2 seconds
```

### Nothing is silently dropped

If a host is found but detail can't be read, it **still appears** in the CSV with a `Status` explaining why:

```
"SQLTEST-SQL01","(port open, detail unavailable)",...,"NOT READ - Access is denied."
"OLDBOX-01","(none detected)",...,"NOT READ - No SQL Server instance registry key present"
"SQL2KBOX","(none detected)",...,"NOT READ - Instance 'SQL2K' listed but no version data in any known registry layout"
```

This is deliberate. A discovery tool that hides what it couldn't reach gives false confidence — you'd never know to go and investigate. Note the third example names the specific instance it couldn't resolve, so you know exactly which host and instance to check by hand. If **no** instances are found at all, the subnet script still writes a valid header-only CSV so downstream tooling doesn't break on a missing file.

### Database storage inventory fields

The inventory script writes a separate CSV with its own columns, since it reports per-instance measurements rather than registry detail:

| Column | Description |
|---|---|
| `ServerName`, `InstanceName`, `FQDN`, `IPAddress` | Inherited from the discovery CSV(s) that identified the host |
| `SQLTcpPort` | The port resolved from the input CSV, SQL Browser, or the instance's own registry configuration, and actually used to connect |
| `SQLRelease`, `SQLVersion` | Inherited from discovery where available |
| `SQLVCoreCount` | Number of **visible, online** SQL Server schedulers — an initial sizing input, not a complete physical-core licensing assessment |
| `MultipleSQLInstances` | `Yes` when more than one instance was found on the same server, otherwise `No` |
| `TotalDatabases` / `DatabasesMeasured` | Databases found vs. databases successfully measured |
| `DataSpaceUsedGB` | Actual occupied data-page space, from `FILEPROPERTY(..., 'SpaceUsed')` |
| `LogSpaceUsedGB` | Actual used transaction-log space, from `DBCC SQLPERF(LOGSPACE)` |
| `TotalSpaceUsedGB` | `DataSpaceUsedGB` + `LogSpaceUsedGB` |
| `ServiceState` | Whether the SQL service was running at scan time |
| `TcpEnabled` / `NamedPipesEnabled` | Which remote protocols were enabled on the instance |
| `Status` | `OK`, `PARTIAL`, `FAILED`, `STOPPED`, or `NOT CONNECTABLE`, with the condition explained |
| `ScanTime` | ISO-8601 timestamp |

`NOT CONNECTABLE` means the instance was running but had both TCP/IP and Named Pipes disabled — enable a remote protocol and restart that instance before rerunning, if its measurements are required. Review the sanitized sample before running:
[`SQL-Database-Storage-Inventory.sample.csv`](./SQL-Database-Storage-Inventory.sample.csv).

Detailed diagnostics — exception and inner-exception details, stack traces, SQL error numbers, operation context, the executing identity, and the input/output paths — are appended to `.\SQL-Database-Storage-Inventory-Diagnostics.txt` by default, only when a warning or error occurs. A fully successful run does not create or modify that file. Console warnings include a diagnostic ID that identifies the matching text-file entry. Error detail is never written to the inventory CSV itself.

---

## 7. Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `The ActiveDirectory module could not be loaded` | RSAT not installed | `Install-WindowsFeature RSAT-AD-PowerShell` (server) or add the RSAT optional feature (Windows client) |
| `Status: NOT READ - Access is denied.` | Not a local admin on the target | Use `-AuthMode Prompt` or `-Credential` with an account that is |
| `Status: NOT READ - The RPC server is unavailable` | Firewall blocking WMI | Allow TCP 135 + the dynamic RPC range, or enable the *Windows Management Instrumentation (WMI-In)* firewall rule |
| `Status: NOT READ - WinRM cannot complete the operation` on an **old** host | Windows 2000/2003 predates WinRM | Nothing to do — the script already retries over DCOM automatically. If DCOM also fails, check that the *Remote Registry* and *Windows Management Instrumentation* services are running on the target |
| Old SQL host found on the network but shows `(port open, detail unavailable)` | Remote WMI disabled or blocked, which is common on legacy servers | Confirm the account is a local admin there; enable remote WMI/DCOM; failing that, the host is still flagged for manual follow-up |
| SQL 2000 / 7.0 instance listed but `SQLVersion` is blank | Instance published a name but no readable build in any known layout | The `Status` column names the instance — check `HKLM\SOFTWARE\Microsoft\MSSQLServer` on that host manually and, if it's a new layout, add the path to `Resolve-SqlInstanceDetail` |
| `PatchLevel` equals `SQLVersion` on an old instance | `Setup\PatchLevel` did not exist before SQL 2008 | Expected — the script falls back to `CurrentVersion`, then `Version`, so the column is never empty |
| Subnet scan finds nothing | UDP 1434 and TCP 1433 both blocked | Check host firewalls; add custom ports via `-SqlPorts` |
| `Prefix /31 is not supported` | No usable host range | Use `/30` or wider |
| `exceeds the MaxHosts safety limit` | Subnet larger than 4,094 hosts | Deliberately raise `-MaxHosts`, or scan in smaller blocks |
| Security warning when running from a UNC path | PowerShell treats network paths as untrusted | Answer `R` (run once), or `Unblock-File .\script.ps1`, or copy locally first |
| Script won't run at all | Execution policy | `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` |
| Inventory instance reported `Status: FAILED` with a SQL login error | Windows admin rights were assumed to be enough, but the identity has no SQL login | Create a login for the account and grant it the permissions in section 3 — local administrator rights do not grant SQL access |
| Inventory instance reported `Status: NOT CONNECTABLE` | TCP/IP and Named Pipes are both disabled on that instance | Enable a remote protocol (SQL Server Configuration Manager) and restart the instance, then rerun |
| Inventory instance reported `Status: PARTIAL` | Some databases measured, others could not be | Check the diagnostic log entry referenced by the console warning's diagnostic ID — usually a missing per-database grant |
| `DBCC SQLPERF(LOGSPACE)` fails for an instance | Missing `VIEW SERVER STATE` (2019 and earlier) or `VIEW SERVER PERFORMANCE STATE` (2022+) | Grant the version-appropriate permission shown in section 3 |
| Inventory can't find an instance that discovery already listed | SQL Browser didn't respond and no port was recorded in the registry | Increase `-InstanceDiscoveryTimeoutMs`, or add the port manually to the input CSV |
| Inventory diagnostics file wasn't created | No warnings or errors occurred | Expected — the file is only created or appended on a warning or error |

---

## 8. Validation

Both scripts were tested end to end in an isolated Azure lab (`sqltest.local`):

| | DC script | Member server script |
|---|---|---|
| Ran on | `sqltest-dc01` | `sqltest-fs01` |
| Discovery method | AD SPN query -> remote registry | 254-address sweep -> remote registry |
| **Elapsed** | **1.3 seconds** | **13.2 seconds** |
| Instances found | 1 | 1 |
| SQL login used | **None** | **None** |

Verified in both cases:
- CSV confirmed absent before the run (`Test-Path` = `False`)
- Script created and populated it (`Test-Path` = `True`)
- All three credential modes exercised: `CurrentUser`, `Prompt`, and explicit `-Credential`
- Error handling proven by running the DC script on a member server without RSAT — it returned a clear, actionable message rather than a stack trace

### Legacy version support

The legacy registry fallbacks were tested against a simulated registry for each SQL era, driving the **actual functions extracted from the shipped scripts** (not a reimplementation). Both scripts scored 8/8:

| Scenario | Result |
|---|---|
| SQL Server 2022 — modern layout | PASS |
| SQL Server 2005 — no `PatchLevel`, falls back to `CurrentVersion` | PASS |
| SQL Server 2000 **named** instance — no `Setup\Version` at all | PASS |
| SQL Server 2000 **default** instance — different hive, absent from `Instance Names` | PASS |
| 32-bit SQL 2000 on 64-bit Windows — `Wow6432Node` redirection | PASS |
| SQL Server 7.0 — pre-2000 | PASS |
| Unknown future major build (18) — must not be discarded | PASS |
| No SQL installed — must return nothing | PASS |

Two further checks were made against the live lab host, since those paths can't be proven by simulation alone:

- **DCOM transport** — forcing `New-CimSession -SessionOption (New-CimSessionOption -Protocol Dcom)` connected and enumerated instances successfully, confirming the pre-WinRM fallback is sound.
- **`MSSQLServer\CurrentVersion`** — the legacy version key was read directly and returned a valid build, confirming the fallback targets a real location rather than a guess.

Both scripts were then re-run end to end against the lab after the legacy changes, with **no regression** — identical output to the pre-change runs.

### Database storage inventory script

The inventory script's correctness relies on a defensive status taxonomy rather than assuming success: every instance and database is reported `OK`, `PARTIAL`, `FAILED`, `STOPPED`, or `NOT CONNECTABLE` with an explanatory detail, and processing continues after an individual server or database failure so one bad instance cannot abort the run. `NOT CONNECTABLE` specifically replaces what would otherwise surface as an unexplained SQL error 26 when both TCP/IP and Named Pipes are disabled. Validate this script in your own non-production environment before relying on its output, in line with the disclaimer at the top of this document — its lab validation history is tracked separately from the two discovery scripts above.

---

## 9. Full script listing

### 9.1 `SQL-Discovery-DomainController.ps1`

```powershell
<#
================================================================================
 SQL Server Discovery - Domain Controller Edition
================================================================================

 Author   : Russell McKee
 Copyright: Copyright (C) 2026 Russell McKee
 SPDX-License-Identifier: GPL-3.0-only
 LinkedIn : https://www.linkedin.com/in/russellwbmckee/
 Version  : 2.1
 Updated  : 25 September 2026

.SYNOPSIS
    Discovers Microsoft SQL Server instances registered in Active Directory and
    reports server name, SQL version, edition, FQDN and IP address.

.DESCRIPTION
    Queries Active Directory Domain Services (ADDS) for objects publishing an
    MSSQLSvc Service Principal Name (SPN). Each discovered host is then queried
    over the remote registry (WMI / StdRegProv) to read the installed instance
    list, product version, patch level and edition.

    This approach requires NO SQL Server login. It needs only local
    administrator rights on the target host, which a Domain Admin already has.
    It is version-agnostic and will report any SQL Server release from 2000
    through 2025 rather than being hard-coded to a single version.

    LEGACY SUPPORT
      Registry layout changed across releases, so several candidate paths are
      probed and the first that yields a build number wins:
        SQL 2005+        ...\Microsoft SQL Server\MSSQL##.INSTANCE\Setup
        SQL 2000 named   ...\Microsoft SQL Server\<InstanceName>\...
        SQL 2000 default ...\Microsoft\MSSQLServer\...   (a different hive)
      Each is also probed under Wow6432Node for 32-bit SQL on 64-bit Windows.
      Version falls back from Setup\Version to MSSQLServer\CurrentVersion, and
      PatchLevel (which did not exist before SQL 2008) falls back to the running
      CurrentVersion. Connections retry over DCOM when WSMan is unavailable,
      which is the norm on the Windows 2000/2003 hosts that run old SQL.

    Data sources:
      - Active Directory  : which servers publish a SQL SPN (the inventory)
      - Remote registry   : version, patch level and edition (the detail)
      - DNS               : fully qualified domain name and IP address

.PARAMETER AuthMode
    CurrentUser : use the identity running the script (default).
    Prompt      : display a secure credential prompt and use those credentials
                  for the remote registry connections.

.PARAMETER Credential
    Optional PSCredential supplied directly. Use this for scheduled or
    automated runs where an interactive prompt is not possible. When supplied
    it takes precedence over -AuthMode.

.PARAMETER OutputCsv
    Destination CSV file. Defaults to the local working directory.

.EXAMPLE
    .\SQL-Discovery-DomainController.ps1
    Runs using the current Windows identity.

.EXAMPLE
    .\SQL-Discovery-DomainController.ps1 -AuthMode Prompt
    Prompts for alternate credentials before scanning.

.EXAMPLE
    # Non-interactive, e.g. a scheduled task
    $c = Get-Credential
    .\SQL-Discovery-DomainController.ps1 -Credential $c

.EXAMPLE
    .\SQL-Discovery-DomainController.ps1 -OutputCsv 'C:\Reports\sql.csv'

.NOTES
    REQUIREMENTS
      - RSAT ActiveDirectory PowerShell module
      - Local administrator rights on the target SQL hosts
      - WMI / RPC (TCP 135 plus dynamic range) open to the targets

    LIMITATION
      Only instances that have registered an MSSQLSvc SPN are visible to this
      method. An instance running under LocalSystem with no SPN will not be
      returned. Use the companion Member Server subnet-scan script to
      cross-check for unregistered instances.

 -------------------------------------------------------------------------------
 DISCLAIMER - PLEASE READ BEFORE RUNNING
 -------------------------------------------------------------------------------
 This script is provided "AS IS", without warranty of any kind, express or
 implied, including but not limited to the warranties of merchantability,
 fitness for a particular purpose and non-infringement.

 In no event shall the author be liable for any claim, damages, loss of data,
 loss of service, or other liability, whether in an action of contract, tort or
 otherwise, arising from, out of, or in connection with the use of this script.

 You run this script entirely at your own risk. It is NOT supported by Microsoft
 and does not constitute official Microsoft guidance.

 ALWAYS TEST IN A NON-PRODUCTION / TEST ENVIRONMENT FIRST and satisfy yourself
 that the behaviour is appropriate before running against production systems.
 Ensure you have written authorisation to query the systems in scope.
 -------------------------------------------------------------------------------
================================================================================
#>

[CmdletBinding()]
param(
    [ValidateSet('CurrentUser', 'Prompt')]
    [string] $AuthMode = 'CurrentUser',

    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [string] $OutputCsv = '.\SQL-Discovery-DomainController.csv',

    [ValidateRange(1, 64)]
    [int] $ThrottleLimit = 16
)

$ErrorActionPreference = 'Stop'
$script:StartTime = Get-Date

# ---------------------------------------------------------------------------
# SQL major build number to product release name.
# Add future releases here; unknown majors are reported rather than discarded.
# ---------------------------------------------------------------------------
$VersionMap = @{
    '6'  = 'SQL Server 6.0 / 6.5'
    '7'  = 'SQL Server 7.0'
    '8'  = 'SQL Server 2000'
    '9'  = 'SQL Server 2005'
    '10' = 'SQL Server 2008 / 2008 R2'
    '11' = 'SQL Server 2012'
    '12' = 'SQL Server 2014'
    '13' = 'SQL Server 2016'
    '14' = 'SQL Server 2017'
    '15' = 'SQL Server 2019'
    '16' = 'SQL Server 2022'
    '17' = 'SQL Server 2025'
}

function Get-SqlRelease {
    param([string] $Version, [hashtable] $Map)

    if ([string]::IsNullOrWhiteSpace($Version)) { return 'Unknown' }
    $major = ($Version -split '\.')[0]
    if ($Map.ContainsKey($major)) { return $Map[$major] }
    return "Unrecognised (major build $major)"
}

# ---------------------------------------------------------------------------
# Credential handling
#
# Precedence:
#   1. -Credential supplied explicitly (best for scheduled/automated runs)
#   2. -AuthMode Prompt  -> interactive secure prompt
#   3. -AuthMode CurrentUser (default) -> identity running the script
# ---------------------------------------------------------------------------
if ($Credential -eq [System.Management.Automation.PSCredential]::Empty) {
    $Credential = $null
}

if ($AuthMode -eq 'Prompt' -and -not $Credential) {
    $Credential = Get-Credential -Message 'Enter credentials with local administrator rights on the target SQL hosts'
    if (-not $Credential) {
        throw 'Credential prompt was cancelled. No scan was performed.'
    }
}

$effectiveAuth = if ($Credential) { "Explicit credential ($($Credential.UserName))" } else { 'CurrentUser' }

Write-Host ''
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host ' SQL Server Discovery - Domain Controller Edition (v2.1)' -ForegroundColor Cyan
Write-Host ' Author: Russell McKee' -ForegroundColor Cyan
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host (" Auth mode  : {0}" -f $effectiveAuth)
Write-Host (" Running as : {0}" -f ([Security.Principal.WindowsIdentity]::GetCurrent().Name))
Write-Host (" Output CSV : {0}" -f $OutputCsv)
Write-Host ''

# ---------------------------------------------------------------------------
# Step 1 - Query Active Directory for MSSQLSvc SPNs
# ---------------------------------------------------------------------------
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    throw "The ActiveDirectory module could not be loaded. Install RSAT-AD-PowerShell. Detail: $($_.Exception.Message)"
}

Write-Host 'Querying Active Directory for MSSQLSvc service principal names...' -ForegroundColor Yellow

try {
    $spnObjects = Get-ADObject -LDAPFilter '(servicePrincipalName=MSSQLSvc/*)' -Properties servicePrincipalName -ErrorAction Stop
}
catch {
    throw "Active Directory query failed: $($_.Exception.Message)"
}

# Deduplicate hosts and retain the SPNs that pointed at each one.
$targets = @{}
foreach ($obj in $spnObjects) {
    foreach ($spn in $obj.servicePrincipalName) {
        if ($spn -notlike 'MSSQLSvc/*') { continue }

        $hostPart = ($spn -split '/', 2)[1]
        $hostName = ($hostPart -split ':')[0]
        $hostName = ($hostName -split '\\')[0]
        if ([string]::IsNullOrWhiteSpace($hostName)) { continue }

        $key = $hostName.ToLowerInvariant()
        if (-not $targets.ContainsKey($key)) {
            $targets[$key] = [System.Collections.Generic.List[string]]::new()
        }
        if (-not $targets[$key].Contains($spn)) { $targets[$key].Add($spn) }
    }
}

Write-Host ("Active Directory returned {0} unique SQL host(s)." -f $targets.Count) -ForegroundColor Green
Write-Host ''

if ($targets.Count -eq 0) {
    Write-Warning 'No MSSQLSvc SPNs were found in Active Directory. Nothing to scan.'
}

# ---------------------------------------------------------------------------
# Step 2 - Interrogate each host via the remote registry
# ---------------------------------------------------------------------------
function Get-RemoteRegString {
    param($CimSession, [string] $SubKey, [string] $ValueName)

    $r = Invoke-CimMethod -CimSession $CimSession -Namespace 'root/cimv2' `
        -ClassName 'StdRegProv' -MethodName 'GetStringValue' `
        -Arguments @{ hDefKey = [uint32] 2147483650; sSubKeyName = $SubKey; sValueName = $ValueName } `
        -ErrorAction SilentlyContinue

    if ($r -and $r.ReturnValue -eq 0) { return $r.sValue }
    return $null
}

function Get-RemoteRegValueNames {
    param($CimSession, [string] $SubKey)

    $r = Invoke-CimMethod -CimSession $CimSession -Namespace 'root/cimv2' `
        -ClassName 'StdRegProv' -MethodName 'EnumValues' `
        -Arguments @{ hDefKey = [uint32] 2147483650; sSubKeyName = $SubKey } `
        -ErrorAction SilentlyContinue

    if ($r -and $r.ReturnValue -eq 0 -and $r.sNames) { return @($r.sNames) }
    return @()
}

function New-DiscoverySession {
    <#
        WinRM did not ship until Windows Server 2008. A SQL Server 2000 or 2005
        host is very likely running Windows 2000/2003, where New-CimSession's
        default WSMan transport will fail. Fall back to DCOM for those.
    #>
    param([string] $ComputerName, $Credential)

    $p = @{ ComputerName = $ComputerName; ErrorAction = 'Stop' }
    if ($Credential) { $p['Credential'] = $Credential }

    try {
        return New-CimSession @p
    }
    catch {
        Write-Verbose ("WSMan failed for {0}, retrying over DCOM: {1}" -f $ComputerName, $_.Exception.Message)
        $p['SessionOption'] = New-CimSessionOption -Protocol Dcom
        return New-CimSession @p
    }
}

function Get-SqlInstanceMap {
    <#
        Returns @{ InstanceName = InstanceKey }.

        'Instance Names\SQL' is the modern location and is checked first, in both
        the native and Wow6432Node views (32-bit SQL on 64-bit Windows was common
        for SQL 2000/2005). If nothing is found there, the SQL Server 2000 default
        instance is probed directly, because it lives in a different hive
        (HKLM\SOFTWARE\Microsoft\MSSQLServer) and does not always publish here.
    #>
    param($CimSession)

    $map = [ordered]@{}

    foreach ($root in @(
            'SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL',
            'SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL'
        )) {
        foreach ($name in (Get-RemoteRegValueNames -CimSession $CimSession -SubKey $root)) {
            if ($map.Contains($name)) { continue }
            $key = Get-RemoteRegString -CimSession $CimSession -SubKey $root -ValueName $name
            if ($key) { $map[$name] = $key }
        }
    }

    if ($map.Count -eq 0) {
        foreach ($legacy in @('SOFTWARE\Microsoft\MSSQLServer',
                              'SOFTWARE\Wow6432Node\Microsoft\MSSQLServer')) {
            $v = Get-RemoteRegString -CimSession $CimSession `
                    -SubKey "$legacy\MSSQLServer\CurrentVersion" -ValueName 'CurrentVersion'
            if (-not $v) {
                $v = Get-RemoteRegString -CimSession $CimSession -SubKey "$legacy\Setup" -ValueName 'SQLPath'
            }
            if ($v) { $map['MSSQLSERVER'] = 'MSSQLSERVER'; break }
        }
    }

    return $map
}

function Resolve-SqlInstanceDetail {
    <#
        Registry layout differs by SQL release, so candidate paths are tried in
        order and the first that yields a build number wins:

          SQL 2005+        ...\Microsoft SQL Server\MSSQL##.INSTANCE\Setup
          SQL 2000 named   ...\Microsoft SQL Server\<InstanceName>\...
          SQL 2000 default ...\Microsoft\MSSQLServer\...          (different hive)

        Version lives in Setup\Version on modern builds and in
        MSSQLServer\CurrentVersion\CurrentVersion on older ones.
        PatchLevel did not exist before SQL 2008, so it falls back to the
        running CurrentVersion, then to Version.
    #>
    param($CimSession, [string] $InstanceKey, [string] $InstanceName)

    $roots = [System.Collections.Generic.List[string]]::new()
    if ($InstanceKey)  { $roots.Add("SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceKey") }
    if ($InstanceName -and $InstanceName -ne $InstanceKey) {
        $roots.Add("SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceName")
    }
    $roots.Add('SOFTWARE\Microsoft\MSSQLServer')

    foreach ($root in @($roots)) {
        foreach ($r in @($root, ($root -replace '^SOFTWARE\\', 'SOFTWARE\Wow6432Node\'))) {

            $setup   = "$r\Setup"
            $current = "$r\MSSQLServer\CurrentVersion"

            $version = Get-RemoteRegString -CimSession $CimSession -SubKey $setup -ValueName 'Version'
            $running = Get-RemoteRegString -CimSession $CimSession -SubKey $current -ValueName 'CurrentVersion'

            if (-not $version) { $version = $running }
            if (-not $version) { continue }

            $patch = Get-RemoteRegString -CimSession $CimSession -SubKey $setup -ValueName 'PatchLevel'
            if (-not $patch) { $patch = $running }
            if (-not $patch) { $patch = $version }

            return [pscustomobject]@{
                Version     = $version
                PatchLevel  = $patch
                Edition     = (Get-RemoteRegString -CimSession $CimSession -SubKey $setup -ValueName 'Edition')
                EditionType = (Get-RemoteRegString -CimSession $CimSession -SubKey $setup -ValueName 'EditionType')
                Collation   = (Get-RemoteRegString -CimSession $CimSession -SubKey $setup -ValueName 'Collation')
            }
        }
    }

    return $null
}

$results = [System.Collections.Generic.List[object]]::new()
$counter = 0

foreach ($target in ($targets.Keys | Sort-Object)) {

    $counter++
    $shortName = ($target -split '\.')[0].ToUpperInvariant()
    Write-Progress -Activity 'SQL Server discovery' `
        -Status ("Querying {0} ({1} of {2})" -f $shortName, $counter, $targets.Count) `
        -PercentComplete (($counter / [Math]::Max($targets.Count, 1)) * 100)

    # --- DNS resolution (non-fatal) -----------------------------------------
    $fqdn = $target
    $ipAddresses = ''
    try {
        $dns = [System.Net.Dns]::GetHostEntry($target)
        $fqdn = $dns.HostName
        $ipAddresses = (
            $dns.AddressList |
            Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
            ForEach-Object { $_.IPAddressToString }
        ) -join '; '
    }
    catch {
        Write-Warning ("DNS resolution failed for {0}: {1}" -f $target, $_.Exception.Message)
    }

    # --- Remote registry ----------------------------------------------------
    $session = $null
    $instancesFound = 0
    $failureReason = $null

    try {
        $session = New-DiscoverySession -ComputerName $shortName -Credential $Credential

        $instanceMap = Get-SqlInstanceMap -CimSession $session

        if ($instanceMap.Count -eq 0) {
            $failureReason = 'No SQL Server instance registry key present'
        }
        else {
            foreach ($instanceName in $instanceMap.Keys) {

                $instanceKey = $instanceMap[$instanceName]
                $detail = Resolve-SqlInstanceDetail -CimSession $session `
                              -InstanceKey $instanceKey -InstanceName $instanceName

                if (-not $detail) {
                    $failureReason = "Instance '$instanceName' listed but no version data in any known registry layout"
                    continue
                }

                $results.Add([pscustomobject][ordered]@{
                    ServerName    = $shortName
                    InstanceName  = $instanceName
                    FQDN          = $fqdn
                    IPAddress     = $ipAddresses
                    SQLRelease    = (Get-SqlRelease -Version $detail.Version -Map $VersionMap)
                    SQLVersion    = $detail.Version
                    PatchLevel    = $detail.PatchLevel
                    SQLEdition    = $detail.Edition
                    EditionType   = $detail.EditionType
                    Collation     = $detail.Collation
                    InstanceKey   = $instanceKey
                    ServicePrincipalName = ($targets[$target] -join '; ')
                    DiscoveredVia = 'Active Directory SPN + Remote Registry'
                    Status        = 'OK'
                    ScanTime      = (Get-Date -Format 's')
                })

                $instancesFound++
            }

            if ($instancesFound -eq 0) { $failureReason = 'Registry key present but no instances enumerated' }
        }
    }
    catch {
        $failureReason = $_.Exception.Message
        Write-Warning ("{0}: {1}" -f $shortName, $failureReason)
    }
    finally {
        if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
    }

    # Always emit a row so an unreachable host is visible rather than silently dropped.
    if ($instancesFound -eq 0) {
        $results.Add([pscustomobject][ordered]@{
            ServerName    = $shortName
            InstanceName  = '(none detected)'
            FQDN          = $fqdn
            IPAddress     = $ipAddresses
            SQLRelease    = ''
            SQLVersion    = ''
            PatchLevel    = ''
            SQLEdition    = ''
            EditionType   = ''
            Collation     = ''
            InstanceKey   = ''
            ServicePrincipalName = ($targets[$target] -join '; ')
            DiscoveredVia = 'Active Directory SPN only'
            Status        = "NOT READ - $failureReason"
            ScanTime      = (Get-Date -Format 's')
        })
    }
}

Write-Progress -Activity 'SQL Server discovery' -Completed

# ---------------------------------------------------------------------------
# Step 3 - Report and export
# ---------------------------------------------------------------------------
$sorted = @($results | Sort-Object ServerName, InstanceName)

Write-Host ''
$sorted | Format-Table ServerName, InstanceName, SQLRelease, SQLVersion, PatchLevel, SQLEdition, FQDN, IPAddress, Status -AutoSize

$okCount = @($sorted | Where-Object { $_.Status -eq 'OK' }).Count

try {
    $outDir = Split-Path -Parent $OutputCsv
    if ($outDir -and -not (Test-Path -LiteralPath $outDir -PathType Container)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }
    $sorted | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host ("CSV exported to : {0}" -f $OutputCsv) -ForegroundColor Green
}
catch {
    Write-Warning "CSV export failed: $($_.Exception.Message)"
}

Write-Host ("Hosts queried   : {0}" -f $targets.Count)
Write-Host ("Instances found : {0}" -f $okCount)
Write-Host ("Elapsed         : {0:N1} seconds" -f ((Get-Date) - $script:StartTime).TotalSeconds)
Write-Host ''
```

### 9.2 `SQL-Discovery-MemberServer.ps1`

```powershell
<#
================================================================================
 SQL Server Discovery - Member Server / Subnet Scan Edition
================================================================================

 Author   : Russell McKee
 Copyright: Copyright (C) 2026 Russell McKee
 SPDX-License-Identifier: GPL-3.0-only
 LinkedIn : https://www.linkedin.com/in/russellwbmckee/
 Version  : 2.1
 Updated  : 25 September 2026

.SYNOPSIS
    Scans an IPv4 subnet for Microsoft SQL Server instances and reports server
    name, SQL version, edition, FQDN and IP address.

.DESCRIPTION
    Designed to run from an ordinary domain member server (no Active Directory
    tooling required). It sweeps every usable address in the supplied CIDR
    subnet using a three-stage approach:

      1. SQL Browser (UDP 1434) - SSRP request returns instance name, version
         and TCP port without any credentials. Catches named instances.
      2. TCP port probe        - for hosts where SQL Browser is disabled or
         firewalled, a direct connect test on the common SQL ports.
      3. Remote registry       - for any responding host, reads the definitive
         instance list, version, patch level and edition via WMI / StdRegProv.

    Stage 3 requires NO SQL Server login - only local administrator rights on
    the target. This is the key difference from the original version, which
    depended on a SQL login and returned nothing when that login did not exist.

    The scan is version-agnostic and will report any SQL Server release from
    2000 through 2025.

    LEGACY SUPPORT
      Registry layout changed across releases, so several candidate paths are
      probed and the first that yields a build number wins:
        SQL 2005+        ...\Microsoft SQL Server\MSSQL##.INSTANCE\Setup
        SQL 2000 named   ...\Microsoft SQL Server\<InstanceName>\...
        SQL 2000 default ...\Microsoft\MSSQLServer\...   (a different hive)
      Each is also probed under Wow6432Node for 32-bit SQL on 64-bit Windows.
      Version falls back from Setup\Version to MSSQLServer\CurrentVersion, and
      PatchLevel (which did not exist before SQL 2008) falls back to the running
      CurrentVersion. Connections retry over DCOM when WSMan is unavailable,
      which is the norm on the Windows 2000/2003 hosts that run old SQL.
      SQL Server 2000 introduced SSRP, so stage 1 finds those hosts too.

    Hosts are scanned in parallel using a runspace pool, so a /24 sweep
    completes in seconds on both Windows PowerShell 5.1 and PowerShell 7.

.PARAMETER TargetSubnet
    IPv4 subnet in CIDR notation, e.g. '10.20.1.0/24'. Prefix /8 to /30.

.PARAMETER AuthMode
    CurrentUser : use the identity running the script (default).
    Prompt      : display a secure credential prompt and use those credentials
                  for the remote registry connections.

.PARAMETER Credential
    Optional PSCredential supplied directly. Use this for scheduled or
    automated runs where an interactive prompt is not possible. When supplied
    it takes precedence over -AuthMode.

.PARAMETER OutputCsv
    Destination CSV file. Defaults to the local working directory.

.EXAMPLE
    .\SQL-Discovery-MemberServer.ps1 -TargetSubnet '10.20.1.0/24'

.EXAMPLE
    .\SQL-Discovery-MemberServer.ps1 -TargetSubnet '192.168.50.0/24' -AuthMode Prompt

.EXAMPLE
    # Non-interactive, e.g. a scheduled task
    $c = Get-Credential
    .\SQL-Discovery-MemberServer.ps1 -TargetSubnet '10.10.0.0/22' -Credential $c

.EXAMPLE
    .\SQL-Discovery-MemberServer.ps1 -TargetSubnet '10.0.0.0/22' -ThrottleLimit 64

.NOTES
    REQUIREMENTS
      - Windows PowerShell 5.1 or PowerShell 7 (both scan in parallel).
      - Outbound UDP 1434 for SQL Browser discovery.
      - Outbound TCP 1433 (and any custom ports) for the port probe.
      - Local administrator rights on targets for version/edition enrichment.

    Hosts that answer on a SQL port but cannot be read via the registry are
    still reported, flagged in the Status column, so nothing is silently lost.

 -------------------------------------------------------------------------------
 DISCLAIMER - PLEASE READ BEFORE RUNNING
 -------------------------------------------------------------------------------
 This script is provided "AS IS", without warranty of any kind, express or
 implied, including but not limited to the warranties of merchantability,
 fitness for a particular purpose and non-infringement.

 In no event shall the author be liable for any claim, damages, loss of data,
 loss of service, or other liability, whether in an action of contract, tort or
 otherwise, arising from, out of, or in connection with the use of this script.

 You run this script entirely at your own risk. It is NOT supported by Microsoft
 and does not constitute official Microsoft guidance.

 ALWAYS TEST IN A NON-PRODUCTION / TEST ENVIRONMENT FIRST and satisfy yourself
 that the behaviour is appropriate before running against production systems.

 NETWORK SCANNING NOTICE: only scan networks you are explicitly authorised to
 scan. Unauthorised scanning may breach policy or law, and may trigger security
 monitoring and alerting.
 -------------------------------------------------------------------------------
================================================================================
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^\d{1,3}(\.\d{1,3}){3}/\d{1,2}$')]
    [string] $TargetSubnet = '10.20.1.0/24',

    [ValidateSet('CurrentUser', 'Prompt')]
    [string] $AuthMode = 'CurrentUser',

    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [string] $OutputCsv = '.\SQL-Discovery-MemberServer.csv',

    [int[]] $SqlPorts = @(1433),

    [ValidateRange(1, 128)]
    [int] $ThrottleLimit = 32,

    [ValidateRange(100, 10000)]
    [int] $TimeoutMs = 700,

    [ValidateRange(1, 65534)]
    [int] $MaxHosts = 4094
)

$ErrorActionPreference = 'Stop'
$script:StartTime = Get-Date

$VersionMap = @{
    '6'  = 'SQL Server 6.0 / 6.5'
    '7'  = 'SQL Server 7.0'
    '8'  = 'SQL Server 2000'
    '9'  = 'SQL Server 2005'
    '10' = 'SQL Server 2008 / 2008 R2'
    '11' = 'SQL Server 2012'
    '12' = 'SQL Server 2014'
    '13' = 'SQL Server 2016'
    '14' = 'SQL Server 2017'
    '15' = 'SQL Server 2019'
    '16' = 'SQL Server 2022'
    '17' = 'SQL Server 2025'
}

# ---------------------------------------------------------------------------
# Expand CIDR to usable host addresses
# ---------------------------------------------------------------------------
function Expand-Subnet {
    param([string] $Cidr, [int] $Limit)

    $parts = $Cidr -split '/'
    $baseIp = [System.Net.IPAddress]::Parse($parts[0])
    $prefix = [int] $parts[1]

    if ($prefix -lt 8 -or $prefix -gt 30) {
        throw "Prefix /$prefix is not supported. Use /8 to /30 (a /31 or /32 has no usable host range)."
    }

    $ipBytes = $baseIp.GetAddressBytes()
    [Array]::Reverse($ipBytes)
    $ipInt = [BitConverter]::ToUInt32($ipBytes, 0)

    $mask = [uint32]([Math]::Pow(2, 32) - [Math]::Pow(2, 32 - $prefix))
    $network = $ipInt -band $mask
    $broadcast = $network -bor (-bnot $mask -band [uint32]::MaxValue)

    $count = $broadcast - $network - 1
    if ($count -gt $Limit) {
        throw "Subnet $Cidr contains $count usable hosts which exceeds the MaxHosts safety limit of $Limit. Raise -MaxHosts if this is intended."
    }

    $list = [System.Collections.Generic.List[string]]::new()
    for ($i = $network + 1; $i -lt $broadcast; $i++) {
        $b = [BitConverter]::GetBytes([uint32] $i)
        [Array]::Reverse($b)
        $list.Add(([System.Net.IPAddress]::new($b)).IPAddressToString)
    }
    return $list
}

# ---------------------------------------------------------------------------
# Credential handling
#
# Precedence:
#   1. -Credential supplied explicitly (best for scheduled/automated runs)
#   2. -AuthMode Prompt  -> interactive secure prompt
#   3. -AuthMode CurrentUser (default) -> identity running the script
# ---------------------------------------------------------------------------
if ($Credential -eq [System.Management.Automation.PSCredential]::Empty) {
    $Credential = $null
}

if ($AuthMode -eq 'Prompt' -and -not $Credential) {
    $Credential = Get-Credential -Message 'Enter credentials with local administrator rights on the target SQL hosts'
    if (-not $Credential) {
        throw 'Credential prompt was cancelled. No scan was performed.'
    }
}

$effectiveAuth = if ($Credential) { "Explicit credential ($($Credential.UserName))" } else { 'CurrentUser' }

Write-Host ''
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host ' SQL Server Discovery - Member Server Subnet Scan (v2.1)' -ForegroundColor Cyan
Write-Host ' Author: Russell McKee' -ForegroundColor Cyan
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host (" Target subnet : {0}" -f $TargetSubnet)
Write-Host (" Auth mode     : {0}" -f $effectiveAuth)
Write-Host (" Running as    : {0}" -f ([Security.Principal.WindowsIdentity]::GetCurrent().Name))
Write-Host (" SQL ports     : {0}" -f ($SqlPorts -join ', '))
Write-Host (" Output CSV    : {0}" -f $OutputCsv)
Write-Host ''

$addresses = Expand-Subnet -Cidr $TargetSubnet -Limit $MaxHosts
Write-Host ("Sweeping {0} usable address(es)..." -f $addresses.Count) -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# Stage 1 + 2 - find live SQL endpoints
#
# A runspace pool is used rather than ForEach-Object -Parallel so that the same
# code path runs at full speed on both Windows PowerShell 5.1 and PowerShell 7.
# ---------------------------------------------------------------------------
$probeScript = {
    param($ip, $ports, $timeout)

    $hit = $null

    # Stage 1 - SQL Browser SSRP on UDP 1434
    $udp = $null
    try {
        $udp = New-Object System.Net.Sockets.UdpClient
        $udp.Client.ReceiveTimeout = $timeout
        $udp.Client.SendTimeout = $timeout
        [void] $udp.Send([byte[]](0x02), 1, $ip, 1434)
        $remote = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        $data = $udp.Receive([ref] $remote)
        if ($data.Length -gt 3) {
            $hit = [pscustomobject]@{
                IP = $ip; Method = 'SQL Browser (UDP 1434)'
                Raw = [Text.Encoding]::ASCII.GetString($data, 3, $data.Length - 3)
            }
        }
    }
    catch { }
    finally { if ($udp) { $udp.Close() } }

    # Stage 2 - TCP probe fallback
    if (-not $hit) {
        foreach ($port in $ports) {
            $tcp = $null
            try {
                $tcp = New-Object System.Net.Sockets.TcpClient
                $async = $tcp.BeginConnect($ip, $port, $null, $null)
                if ($async.AsyncWaitHandle.WaitOne($timeout, $false) -and $tcp.Connected) {
                    $tcp.EndConnect($async)
                    $hit = [pscustomobject]@{ IP = $ip; Method = "TCP port $port"; Raw = '' }
                    break
                }
            }
            catch { }
            finally { if ($tcp) { $tcp.Close() } }
        }
    }

    return $hit
}

$pool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
$pool.Open()
$jobs = [System.Collections.Generic.List[object]]::new()

try {
    foreach ($addr in $addresses) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        [void] $ps.AddScript($probeScript).AddArgument($addr).AddArgument($SqlPorts).AddArgument($TimeoutMs)
        $jobs.Add([pscustomobject]@{ Shell = $ps; Handle = $ps.BeginInvoke() })
    }

    $liveHosts = [System.Collections.Generic.List[object]]::new()
    $done = 0

    foreach ($job in $jobs) {
        try {
            $out = $job.Shell.EndInvoke($job.Handle)
            foreach ($o in $out) { if ($o) { $liveHosts.Add($o) } }
        }
        catch {
            Write-Verbose ("Probe runspace error: {0}" -f $_.Exception.Message)
        }
        finally {
            $job.Shell.Dispose()
        }

        $done++
        if ($done % 16 -eq 0 -or $done -eq $jobs.Count) {
            Write-Progress -Activity 'Subnet sweep' `
                -Status ("{0} of {1} addresses probed" -f $done, $jobs.Count) `
                -PercentComplete (($done / [Math]::Max($jobs.Count, 1)) * 100)
        }
    }
}
finally {
    Write-Progress -Activity 'Subnet sweep' -Completed
    $pool.Close()
    $pool.Dispose()
}

$liveHosts = @($liveHosts | Where-Object { $_ })
Write-Host ("Found {0} responding SQL endpoint(s)." -f $liveHosts.Count) -ForegroundColor Green
Write-Host ''

# ---------------------------------------------------------------------------
# Stage 3 - enrich each hit via remote registry (no SQL login needed)
# ---------------------------------------------------------------------------
function Get-SqlRelease {
    param([string] $Version, [hashtable] $Map)
    if ([string]::IsNullOrWhiteSpace($Version)) { return 'Unknown' }
    $major = ($Version -split '\.')[0]
    if ($Map.ContainsKey($major)) { return $Map[$major] }
    return "Unrecognised (major build $major)"
}

function Get-RemoteRegString {
    param($CimSession, [string] $SubKey, [string] $ValueName)
    $r = Invoke-CimMethod -CimSession $CimSession -Namespace 'root/cimv2' `
        -ClassName 'StdRegProv' -MethodName 'GetStringValue' `
        -Arguments @{ hDefKey = [uint32] 2147483650; sSubKeyName = $SubKey; sValueName = $ValueName } `
        -ErrorAction SilentlyContinue
    if ($r -and $r.ReturnValue -eq 0) { return $r.sValue }
    return $null
}

function Get-RemoteRegValueNames {
    param($CimSession, [string] $SubKey)
    $r = Invoke-CimMethod -CimSession $CimSession -Namespace 'root/cimv2' `
        -ClassName 'StdRegProv' -MethodName 'EnumValues' `
        -Arguments @{ hDefKey = [uint32] 2147483650; sSubKeyName = $SubKey } `
        -ErrorAction SilentlyContinue
    if ($r -and $r.ReturnValue -eq 0 -and $r.sNames) { return @($r.sNames) }
    return @()
}

function New-DiscoverySession {
    <#
        WinRM did not ship until Windows Server 2008. A SQL Server 2000 or 2005
        host is very likely running Windows 2000/2003, where New-CimSession's
        default WSMan transport will fail. Fall back to DCOM for those.
    #>
    param([string] $ComputerName, $Credential)

    $p = @{ ComputerName = $ComputerName; ErrorAction = 'Stop' }
    if ($Credential) { $p['Credential'] = $Credential }

    try {
        return New-CimSession @p
    }
    catch {
        Write-Verbose ("WSMan failed for {0}, retrying over DCOM: {1}" -f $ComputerName, $_.Exception.Message)
        $p['SessionOption'] = New-CimSessionOption -Protocol Dcom
        return New-CimSession @p
    }
}

function Get-SqlInstanceMap {
    <#
        Returns @{ InstanceName = InstanceKey }.

        'Instance Names\SQL' is the modern location and is checked first, in both
        the native and Wow6432Node views (32-bit SQL on 64-bit Windows was common
        for SQL 2000/2005). If nothing is found there, the SQL Server 2000 default
        instance is probed directly, because it lives in a different hive
        (HKLM\SOFTWARE\Microsoft\MSSQLServer) and does not always publish here.
    #>
    param($CimSession)

    $map = [ordered]@{}

    foreach ($root in @(
            'SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL',
            'SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL'
        )) {
        foreach ($name in (Get-RemoteRegValueNames -CimSession $CimSession -SubKey $root)) {
            if ($map.Contains($name)) { continue }
            $key = Get-RemoteRegString -CimSession $CimSession -SubKey $root -ValueName $name
            if ($key) { $map[$name] = $key }
        }
    }

    if ($map.Count -eq 0) {
        foreach ($legacy in @('SOFTWARE\Microsoft\MSSQLServer',
                              'SOFTWARE\Wow6432Node\Microsoft\MSSQLServer')) {
            $v = Get-RemoteRegString -CimSession $CimSession `
                    -SubKey "$legacy\MSSQLServer\CurrentVersion" -ValueName 'CurrentVersion'
            if (-not $v) {
                $v = Get-RemoteRegString -CimSession $CimSession -SubKey "$legacy\Setup" -ValueName 'SQLPath'
            }
            if ($v) { $map['MSSQLSERVER'] = 'MSSQLSERVER'; break }
        }
    }

    return $map
}

function Resolve-SqlInstanceDetail {
    <#
        Registry layout differs by SQL release, so candidate paths are tried in
        order and the first that yields a build number wins:

          SQL 2005+        ...\Microsoft SQL Server\MSSQL##.INSTANCE\Setup
          SQL 2000 named   ...\Microsoft SQL Server\<InstanceName>\...
          SQL 2000 default ...\Microsoft\MSSQLServer\...          (different hive)

        Version lives in Setup\Version on modern builds and in
        MSSQLServer\CurrentVersion\CurrentVersion on older ones.
        PatchLevel did not exist before SQL 2008, so it falls back to the
        running CurrentVersion, then to Version.
    #>
    param($CimSession, [string] $InstanceKey, [string] $InstanceName)

    $roots = [System.Collections.Generic.List[string]]::new()
    if ($InstanceKey)  { $roots.Add("SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceKey") }
    if ($InstanceName -and $InstanceName -ne $InstanceKey) {
        $roots.Add("SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceName")
    }
    $roots.Add('SOFTWARE\Microsoft\MSSQLServer')

    foreach ($root in @($roots)) {
        foreach ($r in @($root, ($root -replace '^SOFTWARE\\', 'SOFTWARE\Wow6432Node\'))) {

            $setup   = "$r\Setup"
            $current = "$r\MSSQLServer\CurrentVersion"

            $version = Get-RemoteRegString -CimSession $CimSession -SubKey $setup -ValueName 'Version'
            $running = Get-RemoteRegString -CimSession $CimSession -SubKey $current -ValueName 'CurrentVersion'

            if (-not $version) { $version = $running }
            if (-not $version) { continue }

            $patch = Get-RemoteRegString -CimSession $CimSession -SubKey $setup -ValueName 'PatchLevel'
            if (-not $patch) { $patch = $running }
            if (-not $patch) { $patch = $version }

            return [pscustomobject]@{
                Version     = $version
                PatchLevel  = $patch
                Edition     = (Get-RemoteRegString -CimSession $CimSession -SubKey $setup -ValueName 'Edition')
                EditionType = (Get-RemoteRegString -CimSession $CimSession -SubKey $setup -ValueName 'EditionType')
                Collation   = (Get-RemoteRegString -CimSession $CimSession -SubKey $setup -ValueName 'Collation')
            }
        }
    }

    return $null
}

$results = [System.Collections.Generic.List[object]]::new()
$n = 0

foreach ($hit in $liveHosts) {

    $n++
    $ip = $hit.IP
    Write-Progress -Activity 'Reading SQL detail' -Status ("{0} ({1} of {2})" -f $ip, $n, $liveHosts.Count) `
        -PercentComplete (($n / [Math]::Max($liveHosts.Count, 1)) * 100)

    # Reverse DNS (non-fatal)
    $fqdn = $ip
    $shortName = $ip
    try {
        $dns = [System.Net.Dns]::GetHostEntry($ip)
        $fqdn = $dns.HostName
        $shortName = ($fqdn -split '\.')[0].ToUpperInvariant()
    }
    catch {
        Write-Verbose ("Reverse DNS failed for {0}: {1}" -f $ip, $_.Exception.Message)
    }

    $session = $null
    $instancesFound = 0
    $failureReason = $null

    try {
        $session = New-DiscoverySession -ComputerName $shortName -Credential $Credential

        $instanceMap = Get-SqlInstanceMap -CimSession $session

        if ($instanceMap.Count -eq 0) {
            $failureReason = 'No SQL Server instance registry key present'
        }
        else {
            foreach ($instanceName in $instanceMap.Keys) {

                $instanceKey = $instanceMap[$instanceName]
                $detail = Resolve-SqlInstanceDetail -CimSession $session `
                              -InstanceKey $instanceKey -InstanceName $instanceName

                if (-not $detail) {
                    $failureReason = "Instance '$instanceName' listed but no version data in any known registry layout"
                    continue
                }

                $results.Add([pscustomobject][ordered]@{
                    ServerName    = $shortName
                    InstanceName  = $instanceName
                    FQDN          = $fqdn
                    IPAddress     = $ip
                    SQLRelease    = (Get-SqlRelease -Version $detail.Version -Map $VersionMap)
                    SQLVersion    = $detail.Version
                    PatchLevel    = $detail.PatchLevel
                    SQLEdition    = $detail.Edition
                    EditionType   = $detail.EditionType
                    Collation     = $detail.Collation
                    InstanceKey   = $instanceKey
                    DetectedBy    = $hit.Method
                    DiscoveredVia = 'Subnet scan + Remote Registry'
                    Status        = 'OK'
                    ScanTime      = (Get-Date -Format 's')
                })

                $instancesFound++
            }
            if ($instancesFound -eq 0 -and -not $failureReason) {
                $failureReason = 'Registry key present but no instances enumerated'
            }
        }
    }
    catch {
        $failureReason = $_.Exception.Message
        Write-Warning ("{0} ({1}): {2}" -f $shortName, $ip, $failureReason)
    }
    finally {
        if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
    }

    # A host that answered on a SQL port is still reported even if the registry
    # could not be read - never silently discard a positive network hit.
    if ($instancesFound -eq 0) {
        $results.Add([pscustomobject][ordered]@{
            ServerName    = $shortName
            InstanceName  = '(port open, detail unavailable)'
            FQDN          = $fqdn
            IPAddress     = $ip
            SQLRelease    = ''
            SQLVersion    = ''
            PatchLevel    = ''
            SQLEdition    = ''
            EditionType   = ''
            Collation     = ''
            InstanceKey   = ''
            DetectedBy    = $hit.Method
            DiscoveredVia = 'Subnet scan only'
            Status        = "NOT READ - $failureReason"
            ScanTime      = (Get-Date -Format 's')
        })
    }
}

Write-Progress -Activity 'Reading SQL detail' -Completed

# ---------------------------------------------------------------------------
# Report and export
# ---------------------------------------------------------------------------
$sorted = @($results | Sort-Object IPAddress, InstanceName)

Write-Host ''
$sorted | Format-Table ServerName, InstanceName, SQLRelease, SQLVersion, PatchLevel, SQLEdition, FQDN, IPAddress, Status -AutoSize

$okCount = @($sorted | Where-Object { $_.Status -eq 'OK' }).Count

try {
    $outDir = Split-Path -Parent $OutputCsv
    if ($outDir -and -not (Test-Path -LiteralPath $outDir -PathType Container)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    if ($sorted.Count -gt 0) {
        $sorted | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
    }
    else {
        # Always leave a valid CSV with headers so downstream tooling does not break.
        '"ServerName","InstanceName","FQDN","IPAddress","SQLRelease","SQLVersion","PatchLevel","SQLEdition","EditionType","Collation","InstanceKey","DetectedBy","DiscoveredVia","Status","ScanTime"' |
            Set-Content -LiteralPath $OutputCsv -Encoding UTF8
        Write-Warning 'No SQL Server endpoints discovered. Check UDP 1434, TCP 1433, routing and host firewalls.'
    }
    Write-Host ("CSV exported to  : {0}" -f $OutputCsv) -ForegroundColor Green
}
catch {
    Write-Warning "CSV export failed: $($_.Exception.Message)"
}

Write-Host ("Addresses swept  : {0}" -f $addresses.Count)
Write-Host ("Endpoints found  : {0}" -f $liveHosts.Count)
Write-Host ("Instances read   : {0}" -f $okCount)
Write-Host ("Elapsed          : {0:N1} seconds" -f ((Get-Date) - $script:StartTime).TotalSeconds)
Write-Host ''
```

### 9.3 `SQL-Database-Storage-Inventory.ps1`

```powershell
<#
================================================================================
 SQL Server Database Storage Inventory
================================================================================

 Author   : Russell McKee
 Copyright: Copyright (C) 2026 Russell McKee
 SPDX-License-Identifier: GPL-3.0-only
 LinkedIn : https://www.linkedin.com/in/russellwbmckee/
 Version  : 2.3
 Updated  : 7 October 2026

.SYNOPSIS
    Expands discovered SQL Server hosts into individual instances, then reports
    SQL-visible vCore counts, database counts, and actual used database storage
    for every instance.

.DESCRIPTION
    Imports one or more CSV files produced by SQL-Discovery-DomainController.ps1
    or SQL-Discovery-MemberServer.ps1 and deduplicates the discovered hosts.
    The inventory then expands each host into instances using SQL Browser plus
    Windows SQL Database Engine services and the remote registry. It connects
    to every resulting instance from a Windows member server.

    By default, connections use Windows integrated authentication with the
    identity running PowerShell. Set -AuthMode Prompt or supply -Credential
    when SQL Server authentication is required. An explicitly supplied
    credential takes precedence over -AuthMode.

    Windows instance discovery has separate authentication controls:
      CurrentUser : use the Windows identity running PowerShell (default).
      Prompt      : securely prompt for a Windows discovery credential.
      BrowserOnly : skip authenticated WMI/registry discovery and use only SQL
                    Browser plus instances already present in the input CSVs.

    Supply -DiscoveryCredential for non-interactive WMI/registry discovery.
    SQL Browser does not require authentication.

    Actual data usage is calculated from FILEPROPERTY(name, 'SpaceUsed') for
    every accessible online database. Actual transaction-log usage is obtained
    from DBCC SQLPERF(LOGSPACE). Results include system databases unless
    -ExcludeSystemDatabases is specified.

    SQLVCoreCount is the number of VISIBLE ONLINE SQL Server schedulers. It
    represents the logical processors currently available to the SQL Database
    Engine and is intended as an initial target-estate sizing input. It is not
    by itself a complete physical-core licensing assessment.

    Each expanded SQL Server instance is inventoried separately.
    MultipleSQLInstances is Yes when more than one instance is found on the
    same server and No otherwise. SQLTcpPort records a port resolved from the
    input CSV, SQL Browser, or the instance's Windows registry configuration.
    Service state and enabled network protocols are exported. An instance with
    both TCP/IP and Named Pipes disabled is reported as NOT CONNECTABLE rather
    than generating a misleading SQL connection failure.

.PARAMETER InputCsv
    One or more discovery CSV paths. By default, the script looks for both
    standard discovery output files in the current directory.

.PARAMETER OutputCsv
    Destination CSV path.

.PARAMETER DiagnosticLog
    Append-only diagnostic text file path. Console errors include a diagnostic
    ID that identifies the corresponding detailed entry. Existing content is
    preserved. The file is created or appended only when an error or warning is
    captured. Error details are not written to the inventory CSV.

.PARAMETER AuthMode
    CurrentUser : use Windows integrated authentication with the identity
                  running PowerShell (default).
    Prompt      : securely prompt for a SQL Server login and password.

.PARAMETER Credential
    Optional SQL Server authentication credential supplied directly. It takes
    precedence over -AuthMode. -SqlCredential remains available as an alias.
    This does not impersonate an alternate Windows account.

.PARAMETER InstanceDiscoveryAuthMode
    CurrentUser : use the current Windows identity for WMI/registry discovery
                  (default).
    Prompt      : securely prompt for a Windows credential.
    BrowserOnly : do not use WMI/registry; use SQL Browser and CSV rows only.

.PARAMETER DiscoveryCredential
    Optional Windows credential for WMI/registry instance discovery. It takes
    precedence over InstanceDiscoveryAuthMode. This credential is never used
    for SQL Server connections.

.PARAMETER InstanceDiscoveryTimeoutMs
    SQL Browser UDP response timeout in milliseconds for each discovered host.

.PARAMETER ExcludeSystemDatabases
    Excludes master, model, msdb and tempdb from the count and storage totals.

.PARAMETER ConnectionTimeoutSeconds
    SQL connection timeout in seconds.

.PARAMETER CommandTimeoutSeconds
    SQL command timeout in seconds for each inventory query.

.PARAMETER Encrypt
    Requests an encrypted SQL connection.

.PARAMETER TrustServerCertificate
    Trusts the server certificate when encryption is enabled. Use only when
    permitted by organizational policy.

.EXAMPLE
    .\SQL-Database-Storage-Inventory.ps1

.EXAMPLE
    .\SQL-Database-Storage-Inventory.ps1 `
        -InputCsv '.\SQL-Discovery-DomainController.csv',
                  '.\SQL-Discovery-MemberServer.csv' `
        -OutputCsv 'C:\Reports\SQL-Database-Storage.csv'

.EXAMPLE
    $credential = Get-Credential
    .\SQL-Database-Storage-Inventory.ps1 -Credential $credential

.EXAMPLE
    .\SQL-Database-Storage-Inventory.ps1 -AuthMode Prompt

.EXAMPLE
    .\SQL-Database-Storage-Inventory.ps1 `
        -InstanceDiscoveryAuthMode Prompt `
        -AuthMode CurrentUser

.EXAMPLE
    $windowsCredential = Get-Credential -Message 'Windows discovery account'
    $sqlCredential = Get-Credential -Message 'SQL login'
    .\SQL-Database-Storage-Inventory.ps1 `
        -DiscoveryCredential $windowsCredential `
        -Credential $sqlCredential

.EXAMPLE
    .\SQL-Database-Storage-Inventory.ps1 `
        -InstanceDiscoveryAuthMode BrowserOnly

.NOTES
    REQUIREMENTS
      - Windows PowerShell 5.1 or PowerShell 7.
      - Network access to every discovered SQL Server endpoint.
      - A login that can connect to each database to be measured.
      - For complete instance expansion, local administrator or equivalent WMI
        and remote-registry access on discovered Windows hosts.
      - WMI / RPC access to targets, or UDP 1434 for BrowserOnly discovery.
      - Permission to run DBCC SQLPERF(LOGSPACE).
      - VIEW SERVER STATE permission to read SQL-visible vCores on SQL Server
        2019 and earlier, or VIEW SERVER PERFORMANCE STATE on SQL Server 2022
        and later.

    Failed, stopped, and non-connectable instances remain in the output CSV.
    Detailed errors and connectivity warnings are appended to the diagnostic
    text file only when they occur.

 -------------------------------------------------------------------------------
 DISCLAIMER - PLEASE READ BEFORE RUNNING
 -------------------------------------------------------------------------------
 This script is provided "AS IS", without warranty of any kind, express or
 implied, including but not limited to the warranties of merchantability,
 fitness for a particular purpose and non-infringement.

 In no event shall the author be liable for any claim, damages, loss of data,
 loss of service, or other liability, whether in an action of contract, tort or
 otherwise, arising from, out of, or in connection with the use of this script.

 You run this script entirely at your own risk. It is NOT supported by Microsoft
 and does not constitute official Microsoft guidance.

 ALWAYS TEST IN A NON-PRODUCTION / TEST ENVIRONMENT FIRST and satisfy yourself
 that the behaviour is appropriate before running against production systems.

 DATABASE QUERY NOTICE: only connect to and query SQL Server instances you are
 explicitly authorised to access. Use a dedicated least-privilege login where
 possible. Unauthorised access may breach policy or law and may trigger security
 monitoring and alerting.
 -------------------------------------------------------------------------------
================================================================================
#>

[CmdletBinding()]
param(
    [string[]] $InputCsv = @(
        '.\SQL-Discovery-DomainController.csv',
        '.\SQL-Discovery-MemberServer.csv'
    ),

    [string] $OutputCsv = '.\SQL-Database-Storage-Inventory.csv',

    [string] $DiagnosticLog = '.\SQL-Database-Storage-Inventory-Diagnostics.txt',

    [ValidateSet('CurrentUser', 'Prompt')]
    [string] $AuthMode = 'CurrentUser',

    [Alias('SqlCredential')]
    [System.Management.Automation.PSCredential]
    $Credential = [System.Management.Automation.PSCredential]::Empty,

    [ValidateSet('CurrentUser', 'Prompt', 'BrowserOnly')]
    [string] $InstanceDiscoveryAuthMode = 'CurrentUser',

    [Alias('WindowsCredential', 'InstanceDiscoveryCredential')]
    [System.Management.Automation.PSCredential]
    $DiscoveryCredential = [System.Management.Automation.PSCredential]::Empty,

    [ValidateRange(100, 10000)]
    [int] $InstanceDiscoveryTimeoutMs = 1000,

    [switch] $ExcludeSystemDatabases,

    [ValidateRange(1, 300)]
    [int] $ConnectionTimeoutSeconds = 10,

    [ValidateRange(1, 3600)]
    [int] $CommandTimeoutSeconds = 30,

    [switch] $Encrypt,

    [switch] $TrustServerCertificate
)

$ErrorActionPreference = 'Stop'
$script:StartTime = Get-Date

if ($Credential -eq [System.Management.Automation.PSCredential]::Empty) {
    $Credential = $null
}

if ($DiscoveryCredential -eq [System.Management.Automation.PSCredential]::Empty) {
    $DiscoveryCredential = $null
}

if ($AuthMode -eq 'Prompt' -and -not $Credential) {
    $Credential = Get-Credential -Message 'Enter a SQL Server login and password for the storage inventory'
    if (-not $Credential) {
        throw 'Credential prompt was cancelled. No inventory was performed.'
    }
}

if ($InstanceDiscoveryAuthMode -eq 'BrowserOnly' -and $DiscoveryCredential) {
    throw 'DiscoveryCredential cannot be used with InstanceDiscoveryAuthMode BrowserOnly.'
}

if ($InstanceDiscoveryAuthMode -eq 'Prompt' -and -not $DiscoveryCredential) {
    $DiscoveryCredential = Get-Credential -Message 'Enter a Windows credential for SQL instance discovery'
    if (-not $DiscoveryCredential) {
        throw 'Windows discovery credential prompt was cancelled. No inventory was performed.'
    }
}

$VersionMap = @{
    '6'  = 'SQL Server 6.0 / 6.5'
    '7'  = 'SQL Server 7.0'
    '8'  = 'SQL Server 2000'
    '9'  = 'SQL Server 2005'
    '10' = 'SQL Server 2008 / 2008 R2'
    '11' = 'SQL Server 2012'
    '12' = 'SQL Server 2014'
    '13' = 'SQL Server 2016'
    '14' = 'SQL Server 2017'
    '15' = 'SQL Server 2019'
    '16' = 'SQL Server 2022'
    '17' = 'SQL Server 2025'
}

function Get-PropertyValue {
    param(
        [psobject] $InputObject,
        [string] $Name
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if ($property) { return [string] $property.Value }
    return ''
}

function Get-SqlRelease {
    param([string] $Version)

    if ([string]::IsNullOrWhiteSpace($Version)) { return 'Unknown' }
    $major = ($Version -split '\.')[0]
    if ($VersionMap.ContainsKey($major)) { return $VersionMap[$major] }
    return "Unrecognised (major build $major)"
}

function Get-DiscoveryHostName {
    param([psobject] $DiscoveryRow)

    $hostName = Get-PropertyValue -InputObject $DiscoveryRow -Name 'FQDN'
    if ([string]::IsNullOrWhiteSpace($hostName)) {
        $hostName = Get-PropertyValue -InputObject $DiscoveryRow -Name 'ServerName'
    }
    if ([string]::IsNullOrWhiteSpace($hostName)) {
        $hostName = Get-PropertyValue -InputObject $DiscoveryRow -Name 'IPAddress'
    }
    return $hostName
}

function Get-SqlDataSource {
    param([psobject] $DiscoveryRow)

    $hostName = Get-DiscoveryHostName -DiscoveryRow $DiscoveryRow
    $instanceName = Get-PropertyValue -InputObject $DiscoveryRow -Name 'InstanceName'
    $detectedBy = Get-PropertyValue -InputObject $DiscoveryRow -Name 'DetectedBy'
    $spns = Get-PropertyValue -InputObject $DiscoveryRow -Name 'ServicePrincipalName'

    if ([string]::IsNullOrWhiteSpace($hostName) -or
        [string]::IsNullOrWhiteSpace($instanceName) -or
        $instanceName -like '(*') {
        return $null
    }

    if ($instanceName -ne 'MSSQLSERVER') {
        return "$hostName\$instanceName"
    }

    if ($detectedBy -match '(?i)\bTCP port\s+(?<Port>\d{1,5})\b') {
        return "tcp:$hostName,$($Matches.Port)"
    }

    foreach ($spn in ($spns -split '\s*;\s*')) {
        if ($spn -match '(?i)^MSSQLSvc/[^:]+:(?<Port>\d{1,5})$') {
            return "tcp:$hostName,$($Matches.Port)"
        }
    }

    return $hostName
}

function Get-DiscoveryTcpPort {
    param([psobject] $DiscoveryRow)

    $detectedBy = Get-PropertyValue -InputObject $DiscoveryRow -Name 'DetectedBy'
    if ($detectedBy -match '(?i)\bTCP port\s+(?<Port>\d{1,5})\b') {
        return Get-ValidSqlTcpPort -Value $Matches.Port
    }

    $spns = Get-PropertyValue -InputObject $DiscoveryRow -Name 'ServicePrincipalName'
    foreach ($spn in ($spns -split '\s*;\s*')) {
        if ($spn -match '(?i)^MSSQLSvc/[^:]+:(?<Port>\d{1,5})$') {
            return Get-ValidSqlTcpPort -Value $Matches.Port
        }
    }

    return ''
}

function Get-ServerKey {
    param([psobject] $DiscoveryRow)

    $serverName = Get-PropertyValue -InputObject $DiscoveryRow -Name 'ServerName'
    $fqdn = Get-PropertyValue -InputObject $DiscoveryRow -Name 'FQDN'
    $ipAddress = Get-PropertyValue -InputObject $DiscoveryRow -Name 'IPAddress'

    $hostIdentity = $serverName
    if ([string]::IsNullOrWhiteSpace($hostIdentity)) { $hostIdentity = $fqdn }
    if ([string]::IsNullOrWhiteSpace($hostIdentity)) { $hostIdentity = $ipAddress }

    if ([string]::IsNullOrWhiteSpace($hostIdentity)) { return $null }
    return $hostIdentity.Trim()
}

function Get-InstanceKey {
    param([psobject] $DiscoveryRow)

    $serverKey = Get-ServerKey -DiscoveryRow $DiscoveryRow
    $instanceName = Get-PropertyValue -InputObject $DiscoveryRow -Name 'InstanceName'

    return ('{0}|{1}' -f $serverKey, $instanceName.Trim())
}

function New-ExpandedSqlDataSource {
    param(
        [psobject] $DiscoveryRow,
        [string] $InstanceName,
        [string] $TcpPort
    )

    $hostName = Get-DiscoveryHostName -DiscoveryRow $DiscoveryRow
    if ([string]::IsNullOrWhiteSpace($hostName)) { return $null }

    if (-not [string]::IsNullOrWhiteSpace($TcpPort)) {
        return "tcp:$hostName,$TcpPort"
    }

    if ($InstanceName -eq 'MSSQLSERVER') { return $hostName }
    return "$hostName\$InstanceName"
}

function ConvertFrom-SqlBrowserResponse {
    param([string] $RawResponse)

    $instances = [ordered]@{}
    if ([string]::IsNullOrWhiteSpace($RawResponse)) { return @() }

    foreach ($block in ($RawResponse -split ';;')) {
        $tokens = @($block.Trim(';') -split ';')
        $values = @{}
        for ($i = 0; $i + 1 -lt $tokens.Count; $i += 2) {
            if (-not [string]::IsNullOrWhiteSpace($tokens[$i])) {
                $values[$tokens[$i]] = $tokens[$i + 1]
            }
        }

        $instanceName = [string] $values['InstanceName']
        if ([string]::IsNullOrWhiteSpace($instanceName) -or $instances.Contains($instanceName)) {
            continue
        }

        $tcpPort = [string] $values['tcp']
        if ($tcpPort -notmatch '^\d{1,5}$' -or [int] $tcpPort -gt 65535) {
            $tcpPort = ''
        }

        $instances[$instanceName] = [pscustomobject]@{
            InstanceName = $instanceName
            TcpPort      = $tcpPort
        }
    }

    return @($instances.Values)
}

function Get-SqlBrowserInstances {
    param(
        [string] $HostName,
        [int] $TimeoutMs
    )

    $udp = $null
    try {
        $udp = [System.Net.Sockets.UdpClient]::new()
        $udp.Client.ReceiveTimeout = $TimeoutMs
        $udp.Client.SendTimeout = $TimeoutMs
        [void] $udp.Send([byte[]](0x02), 1, $HostName, 1434)
        $remote = [System.Net.IPEndPoint]::new([System.Net.IPAddress]::Any, 0)
        $data = $udp.Receive([ref] $remote)
        if ($data.Length -le 3) { return @() }

        $rawResponse = [Text.Encoding]::ASCII.GetString($data, 3, $data.Length - 3)
        return @(ConvertFrom-SqlBrowserResponse -RawResponse $rawResponse)
    }
    catch {
        return @()
    }
    finally {
        if ($udp) { $udp.Close() }
    }
}

function Get-RemoteRegString {
    param($CimSession, [string] $SubKey, [string] $ValueName)

    $result = Invoke-CimMethod -CimSession $CimSession -Namespace 'root/cimv2' `
        -ClassName 'StdRegProv' -MethodName 'GetStringValue' `
        -Arguments @{ hDefKey = [uint32] 2147483650; sSubKeyName = $SubKey; sValueName = $ValueName } `
        -ErrorAction Stop
    if ($result -and $result.ReturnValue -eq 0) { return $result.sValue }
    return $null
}

function Get-RemoteRegValueNames {
    param($CimSession, [string] $SubKey)

    $result = Invoke-CimMethod -CimSession $CimSession -Namespace 'root/cimv2' `
        -ClassName 'StdRegProv' -MethodName 'EnumValues' `
        -Arguments @{ hDefKey = [uint32] 2147483650; sSubKeyName = $SubKey } `
        -ErrorAction Stop
    if ($result -and $result.ReturnValue -eq 0 -and $result.sNames) {
        return @($result.sNames)
    }
    return @()
}

function Get-RemoteRegSubKeyNames {
    param($CimSession, [string] $SubKey)

    $result = Invoke-CimMethod -CimSession $CimSession -Namespace 'root/cimv2' `
        -ClassName 'StdRegProv' -MethodName 'EnumKey' `
        -Arguments @{ hDefKey = [uint32] 2147483650; sSubKeyName = $SubKey } `
        -ErrorAction Stop
    if ($result -and $result.ReturnValue -eq 0 -and $result.sNames) {
        return @($result.sNames)
    }
    return @()
}

function Get-RemoteRegDword {
    param($CimSession, [string] $SubKey, [string] $ValueName)

    $result = Invoke-CimMethod -CimSession $CimSession -Namespace 'root/cimv2' `
        -ClassName 'StdRegProv' -MethodName 'GetDWORDValue' `
        -Arguments @{ hDefKey = [uint32] 2147483650; sSubKeyName = $SubKey; sValueName = $ValueName } `
        -ErrorAction Stop
    if ($result -and $result.ReturnValue -eq 0) {
        return [int] $result.uValue
    }
    return $null
}

function Get-ValidSqlTcpPort {
    param([string] $Value)

    foreach ($candidate in ($Value -split '[,;\s]+')) {
        if ($candidate -match '^\d{1,5}$') {
            $port = [int] $candidate
            if ($port -ge 1 -and $port -le 65535) {
                return [string] $port
            }
        }
    }
    return ''
}

function Get-SqlInstanceTcpPort {
    param(
        $CimSession,
        [string] $InstanceKey
    )

    if ([string]::IsNullOrWhiteSpace($InstanceKey)) { return '' }

    foreach ($root in @(
            "SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceKey\MSSQLServer\SuperSocketNetLib\Tcp",
            "SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\$InstanceKey\MSSQLServer\SuperSocketNetLib\Tcp"
        )) {
        $networkKeys = [System.Collections.Generic.List[string]]::new()
        $networkKeys.Add("$root\IPAll")
        $networkKeys.Add($root)
        foreach ($subKey in (Get-RemoteRegSubKeyNames -CimSession $CimSession -SubKey $root)) {
            if ($subKey -ne 'IPAll') {
                $networkKeys.Add("$root\$subKey")
            }
        }

        foreach ($networkKey in $networkKeys) {
            foreach ($valueName in 'TcpPort', 'TcpDynamicPorts') {
                $port = Get-ValidSqlTcpPort -Value (
                    Get-RemoteRegString -CimSession $CimSession `
                        -SubKey $networkKey -ValueName $valueName
                )
                if ($port) { return $port }
            }
        }
    }

    return ''
}

function Get-SqlInstanceProtocolConfiguration {
    param(
        $CimSession,
        [string] $InstanceKey,
        [string] $InstanceName
    )

    $configuration = [ordered]@{
        TcpEnabled        = $null
        NamedPipesEnabled = $null
        NamedPipe         = ''
    }

    if ([string]::IsNullOrWhiteSpace($InstanceKey)) {
        return [pscustomobject] $configuration
    }

    foreach ($root in @(
            "SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceKey\MSSQLServer\SuperSocketNetLib",
            "SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\$InstanceKey\MSSQLServer\SuperSocketNetLib"
        )) {
        if ($null -eq $configuration.TcpEnabled) {
            $configuration.TcpEnabled = Get-RemoteRegDword -CimSession $CimSession `
                -SubKey "$root\Tcp" -ValueName 'Enabled'
        }
        if ($null -eq $configuration.NamedPipesEnabled) {
            $configuration.NamedPipesEnabled = Get-RemoteRegDword -CimSession $CimSession `
                -SubKey "$root\Np" -ValueName 'Enabled'
        }
        if ([string]::IsNullOrWhiteSpace($configuration.NamedPipe)) {
            $configuration.NamedPipe = [string] (
                Get-RemoteRegString -CimSession $CimSession `
                    -SubKey "$root\Np" -ValueName 'PipeName'
            )
        }
    }

    if ($configuration.NamedPipesEnabled -eq 1 -and
        [string]::IsNullOrWhiteSpace($configuration.NamedPipe)) {
        $configuration.NamedPipe = if ($InstanceName -eq 'MSSQLSERVER') {
            '\\.\pipe\sql\query'
        }
        else {
            "\\.\pipe\MSSQL`$$InstanceName\sql\query"
        }
    }

    return [pscustomobject] $configuration
}

function Get-SqlInstanceKeyFromServicePath {
    param([string] $PathName)

    if ($PathName -match '(?i)\\(?<InstanceKey>MSSQL\d+\.[^\\"]+)\\MSSQL\\Binn\\sqlservr\.exe') {
        return $Matches.InstanceKey
    }
    return ''
}

function New-RemoteNamedPipeDataSource {
    param(
        [psobject] $DiscoveryRow,
        [string] $PipeName
    )

    if ([string]::IsNullOrWhiteSpace($PipeName)) { return $null }
    $hostName = Get-DiscoveryHostName -DiscoveryRow $DiscoveryRow
    if ([string]::IsNullOrWhiteSpace($hostName)) { return $null }

    $pipePath = $PipeName.Trim()
    if ($pipePath -match '^\\\\\.\\(?<Path>.+)$') {
        $pipePath = $Matches.Path
    }
    elseif ($pipePath -match '^\\\\[^\\]+\\(?<Path>.+)$') {
        $pipePath = $Matches.Path
    }
    elseif ($pipePath -notmatch '(?i)^pipe\\') {
        $pipePath = "pipe\$pipePath"
    }

    return "np:\\$hostName\$pipePath"
}

function Get-SqlServiceListeningTcpPorts {
    param(
        $CimSession,
        [uint32] $ProcessId
    )

    if ($ProcessId -eq 0) { return @() }

    try {
        $connections = Get-CimInstance -CimSession $CimSession `
            -Namespace 'root/StandardCimv2' -ClassName 'MSFT_NetTCPConnection' `
            -Filter "OwningProcess = $ProcessId AND State = 2" -ErrorAction Stop
    }
    catch {
        return @()
    }

    $ports = @(
        $connections |
        Where-Object { $_.LocalAddress -notin @('127.0.0.1', '::1') } |
        ForEach-Object { Get-ValidSqlTcpPort -Value ([string] $_.LocalPort) } |
        Where-Object { $_ } |
        Group-Object |
        Sort-Object @{ Expression = 'Count'; Descending = $true },
                    @{ Expression = { [int] $_.Name }; Ascending = $true } |
        ForEach-Object { $_.Name }
    )
    return $ports
}

function New-InstanceDiscoverySession {
    param(
        [string] $ComputerName,
        [System.Management.Automation.PSCredential] $WindowsCredential
    )

    $parameters = @{ ComputerName = $ComputerName; ErrorAction = 'Stop' }
    if ($WindowsCredential) { $parameters['Credential'] = $WindowsCredential }

    try {
        return New-CimSession @parameters
    }
    catch {
        $parameters['SessionOption'] = New-CimSessionOption -Protocol Dcom
        return New-CimSession @parameters
    }
}

function Get-WindowsSqlInstances {
    param($CimSession)

    $instances = [ordered]@{}
    foreach ($root in @(
            'SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL',
            'SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL'
        )) {
        foreach ($name in (Get-RemoteRegValueNames -CimSession $CimSession -SubKey $root)) {
            if ($instances.Contains($name)) { continue }
            $instanceKey = Get-RemoteRegString -CimSession $CimSession -SubKey $root -ValueName $name
            if ($instanceKey) {
                $ports = [System.Collections.Generic.List[string]]::new()
                $namedPipes = [System.Collections.Generic.List[string]]::new()
                $registryPort = Get-SqlInstanceTcpPort -CimSession $CimSession `
                    -InstanceKey $instanceKey
                if ($registryPort) { $ports.Add($registryPort) }
                $protocols = Get-SqlInstanceProtocolConfiguration -CimSession $CimSession `
                    -InstanceKey $instanceKey -InstanceName $name
                if ($protocols.NamedPipesEnabled -eq 1 -and $protocols.NamedPipe) {
                    $namedPipes.Add($protocols.NamedPipe)
                }
                $instances[$name] = [pscustomobject]@{
                    InstanceKey      = $instanceKey
                    TcpPorts         = $ports
                    NamedPipes       = $namedPipes
                    ConfiguredNamedPipe = $protocols.NamedPipe
                    TcpEnabled       = $protocols.TcpEnabled
                    NamedPipesEnabled = $protocols.NamedPipesEnabled
                    ServiceState     = ''
                    ServiceStartMode = ''
                }
            }
        }
    }

    $services = Get-CimInstance -CimSession $CimSession -ClassName Win32_Service `
        -Filter "Name = 'MSSQLSERVER' OR Name LIKE 'MSSQL$%'" -ErrorAction Stop
    foreach ($service in @($services)) {
        $instanceName = if ($service.Name -eq 'MSSQLSERVER') {
            'MSSQLSERVER'
        }
        elseif ($service.Name -like 'MSSQL$*') {
            $service.Name.Substring(6)
        }
        else {
            continue
        }

        $serviceInstanceKey = Get-SqlInstanceKeyFromServicePath -PathName $service.PathName
        if (-not $instances.Contains($instanceName)) {
            $ports = [System.Collections.Generic.List[string]]::new()
            $namedPipes = [System.Collections.Generic.List[string]]::new()
            $protocols = Get-SqlInstanceProtocolConfiguration -CimSession $CimSession `
                -InstanceKey $serviceInstanceKey -InstanceName $instanceName
            $registryPort = Get-SqlInstanceTcpPort -CimSession $CimSession `
                -InstanceKey $serviceInstanceKey
            if ($registryPort) { $ports.Add($registryPort) }
            if ($protocols.NamedPipesEnabled -eq 1 -and $protocols.NamedPipe) {
                $namedPipes.Add($protocols.NamedPipe)
            }
            $instances[$instanceName] = [pscustomobject]@{
                InstanceKey      = $serviceInstanceKey
                TcpPorts         = $ports
                NamedPipes       = $namedPipes
                ConfiguredNamedPipe = $protocols.NamedPipe
                TcpEnabled       = $protocols.TcpEnabled
                NamedPipesEnabled = $protocols.NamedPipesEnabled
                ServiceState     = ''
                ServiceStartMode = ''
            }
        }

        $instances[$instanceName].ServiceState = [string] $service.State
        $instances[$instanceName].ServiceStartMode = [string] $service.StartMode
        foreach ($port in @(Get-SqlServiceListeningTcpPorts -CimSession $CimSession `
                -ProcessId ([uint32] $service.ProcessId))) {
            if (-not $instances[$instanceName].TcpPorts.Contains($port)) {
                $instances[$instanceName].TcpPorts.Add($port)
            }
        }
    }

    return $instances
}

function New-SqlConnectionString {
    param(
        [string] $DataSource,
        [string] $Database
    )

    $builder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new()
    $builder['Data Source'] = $DataSource
    $builder['Initial Catalog'] = $Database
    $builder['Connect Timeout'] = $ConnectionTimeoutSeconds
    $builder['Application Name'] = 'SQL Database Storage Inventory'
    $builder['Encrypt'] = [bool] $Encrypt
    $builder['TrustServerCertificate'] = [bool] $TrustServerCertificate

    if ($Credential) {
        $builder['Integrated Security'] = $false
        $builder['User ID'] = $Credential.UserName
        $builder['Password'] = $Credential.GetNetworkCredential().Password
    }
    else {
        $builder['Integrated Security'] = $true
    }

    return $builder.ConnectionString
}

function Invoke-SqlDataTable {
    param(
        [System.Data.SqlClient.SqlConnection] $Connection,
        [string] $Query
    )

    $command = $Connection.CreateCommand()
    $command.CommandText = $Query
    $command.CommandTimeout = $CommandTimeoutSeconds

    $table = [System.Data.DataTable]::new()
    $reader = $null
    try {
        $reader = $command.ExecuteReader()
        $table.Load($reader)
    }
    finally {
        if ($reader) { $reader.Dispose() }
        $command.Dispose()
    }

    return ,$table
}

function Format-ErrorRecord {
    param([System.Management.Automation.ErrorRecord] $ErrorRecord)

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Message: $($ErrorRecord.Exception.Message)")
    $lines.Add("Exception type: $($ErrorRecord.Exception.GetType().FullName)")
    $lines.Add("Fully qualified error ID: $($ErrorRecord.FullyQualifiedErrorId)")
    $lines.Add("Category: $($ErrorRecord.CategoryInfo)")

    if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.PositionMessage) {
        $lines.Add("Invocation: $($ErrorRecord.InvocationInfo.PositionMessage.Trim())")
    }
    if ($ErrorRecord.ScriptStackTrace) {
        $lines.Add("PowerShell stack trace: $($ErrorRecord.ScriptStackTrace)")
    }

    $exception = $ErrorRecord.Exception
    $depth = 0
    while ($exception) {
        $prefix = if ($depth -eq 0) { 'Exception' } else { "Inner exception $depth" }
        $lines.Add("$prefix type: $($exception.GetType().FullName)")
        $lines.Add("$prefix message: $($exception.Message)")
        if ($exception.StackTrace) {
            $lines.Add("$prefix .NET stack trace: $($exception.StackTrace)")
        }

        if ($exception -is [System.Data.SqlClient.SqlException]) {
            foreach ($sqlError in $exception.Errors) {
                $lines.Add(
                    "SQL error: Number=$($sqlError.Number); State=$($sqlError.State); " +
                    "Class=$($sqlError.Class); Server=$($sqlError.Server); " +
                    "Procedure=$($sqlError.Procedure); Line=$($sqlError.LineNumber); " +
                    "Message=$($sqlError.Message)"
                )
            }
        }

        $exception = $exception.InnerException
        $depth++
    }

    return $lines -join [Environment]::NewLine
}

function Write-Diagnostic {
    param(
        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string] $Level,
        [string] $Context,
        [string] $Message,
        [System.Management.Automation.ErrorRecord] $ErrorRecord
    )

    $diagnosticId = [Guid]::NewGuid().ToString('N').Substring(0, 12)
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('--------------------------------------------------------------------------------')
    $lines.Add("Timestamp: $((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.fffK'))")
    $lines.Add("Diagnostic ID: $diagnosticId")
    $lines.Add("Level: $Level")
    $lines.Add("Context: $Context")
    $lines.Add("PowerShell: $($PSVersionTable.PSVersion)")
    $lines.Add("Run as: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)")
    $lines.Add("Input CSV: $($InputCsv -join '; ')")
    $lines.Add("Output CSV: $OutputCsv")
    if ($Message) {
        $lines.Add("Summary: $Message")
    }
    if ($ErrorRecord) {
        $lines.Add((Format-ErrorRecord -ErrorRecord $ErrorRecord))
    }

    $diagnosticDirectory = Split-Path -Parent $script:DiagnosticLogPath
    if ($diagnosticDirectory -and -not (Test-Path -LiteralPath $diagnosticDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $diagnosticDirectory -Force | Out-Null
    }
    Add-Content -LiteralPath $script:DiagnosticLogPath -Value ($lines -join [Environment]::NewLine) -Encoding UTF8
    return $diagnosticId
}

try {
    $script:DiagnosticLogPath = [System.IO.Path]::GetFullPath($DiagnosticLog)
}
catch {
    throw "Diagnostic log path '$DiagnosticLog' is invalid: $($_.Exception.Message)"
}

function Invoke-StorageInventory {
$inputRows = [System.Collections.Generic.List[object]]::new()
$loadedFiles = [System.Collections.Generic.List[string]]::new()

foreach ($path in $InputCsv) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $message = "Discovery CSV not found; skipping: $path"
        $diagnosticId = Write-Diagnostic -Level WARNING -Context 'Input validation' -Message $message
        Write-Warning "$message Diagnostic ID: $diagnosticId"
        continue
    }

    try {
        foreach ($row in @(Import-Csv -LiteralPath $path)) {
            $inputRows.Add($row)
        }
        $loadedFiles.Add((Resolve-Path -LiteralPath $path).Path)
    }
    catch {
        $diagnosticId = Write-Diagnostic -Level ERROR -Context "Importing discovery CSV '$path'" `
            -Message 'The discovery CSV could not be imported and was skipped.' -ErrorRecord $_
        Write-Warning "Could not import '$path'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
    }
}

if ($loadedFiles.Count -eq 0) {
    throw 'None of the supplied discovery CSV files could be found.'
}

$servers = @{}
$skippedRows = 0

foreach ($row in $inputRows) {
    $serverKey = Get-ServerKey -DiscoveryRow $row
    if ([string]::IsNullOrWhiteSpace($serverKey)) {
        $skippedRows++
        continue
    }

    if (-not $servers.ContainsKey($serverKey)) {
        $servers[$serverKey] = [pscustomobject]@{
            ServerKey = $serverKey
            Rows      = [System.Collections.Generic.List[object]]::new()
        }
    }
    $servers[$serverKey].Rows.Add($row)
}

if ($servers.Count -eq 0) {
    throw 'No SQL Server hosts were present in the discovery CSV files.'
}

$instances = @{}
foreach ($server in ($servers.Values | Sort-Object ServerKey)) {
    $baseRow = $server.Rows[0]
    $hostName = Get-DiscoveryHostName -DiscoveryRow $baseRow
    $candidates = [ordered]@{}

    foreach ($row in $server.Rows) {
        $instanceName = Get-PropertyValue -InputObject $row -Name 'InstanceName'
        if ([string]::IsNullOrWhiteSpace($instanceName) -or $instanceName -like '(*') {
            continue
        }

        if (-not $candidates.Contains($instanceName)) {
            $candidates[$instanceName] = [pscustomobject]@{
                DiscoveryRow = $row
                InstanceName = $instanceName
                TcpPorts     = [System.Collections.Generic.List[string]]::new()
                NamedPipes   = [System.Collections.Generic.List[string]]::new()
                InstanceKey  = ''
                ConfiguredNamedPipe = ''
                TcpEnabled   = $null
                NamedPipesEnabled = $null
                ServiceState = ''
                ServiceStartMode = ''
                Sources      = [System.Collections.Generic.List[string]]::new()
            }
        }
        $csvPort = Get-DiscoveryTcpPort -DiscoveryRow $row
        if ($csvPort -and -not $candidates[$instanceName].TcpPorts.Contains($csvPort)) {
            $candidates[$instanceName].TcpPorts.Add($csvPort)
        }
        if (-not $candidates[$instanceName].Sources.Contains('Input CSV')) {
            $candidates[$instanceName].Sources.Add('Input CSV')
        }
    }

    foreach ($browserInstance in @(Get-SqlBrowserInstances -HostName $hostName `
            -TimeoutMs $InstanceDiscoveryTimeoutMs)) {
        $instanceName = $browserInstance.InstanceName
        if (-not $candidates.Contains($instanceName)) {
            $candidates[$instanceName] = [pscustomobject]@{
                DiscoveryRow = $baseRow
                InstanceName = $instanceName
                TcpPorts     = [System.Collections.Generic.List[string]]::new()
                NamedPipes   = [System.Collections.Generic.List[string]]::new()
                InstanceKey  = ''
                ConfiguredNamedPipe = ''
                TcpEnabled   = $null
                NamedPipesEnabled = $null
                ServiceState = ''
                ServiceStartMode = ''
                Sources      = [System.Collections.Generic.List[string]]::new()
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($browserInstance.TcpPort) -and
            -not $candidates[$instanceName].TcpPorts.Contains($browserInstance.TcpPort)) {
            $candidates[$instanceName].TcpPorts.Add($browserInstance.TcpPort)
        }
        if (-not $candidates[$instanceName].Sources.Contains('SQL Browser')) {
            $candidates[$instanceName].Sources.Add('SQL Browser')
        }
    }

    if ($InstanceDiscoveryAuthMode -ne 'BrowserOnly') {
        $session = $null
        try {
            $session = New-InstanceDiscoverySession -ComputerName $hostName `
                -WindowsCredential $DiscoveryCredential
            $windowsInstances = Get-WindowsSqlInstances -CimSession $session
            foreach ($instanceName in $windowsInstances.Keys) {
                $windowsInstance = $windowsInstances[$instanceName]
                if (-not $candidates.Contains($instanceName)) {
                    $candidates[$instanceName] = [pscustomobject]@{
                        DiscoveryRow = $baseRow
                        InstanceName = $instanceName
                        TcpPorts     = [System.Collections.Generic.List[string]]::new()
                        NamedPipes   = [System.Collections.Generic.List[string]]::new()
                        InstanceKey  = ''
                        ConfiguredNamedPipe = ''
                        TcpEnabled   = $null
                        NamedPipesEnabled = $null
                        ServiceState = ''
                        ServiceStartMode = ''
                        Sources      = [System.Collections.Generic.List[string]]::new()
                    }
                }
                foreach ($port in $windowsInstance.TcpPorts) {
                    if (-not $candidates[$instanceName].TcpPorts.Contains($port)) {
                        $candidates[$instanceName].TcpPorts.Add($port)
                    }
                }
                foreach ($pipeName in $windowsInstance.NamedPipes) {
                    if (-not $candidates[$instanceName].NamedPipes.Contains($pipeName)) {
                        $candidates[$instanceName].NamedPipes.Add($pipeName)
                    }
                }
                $candidates[$instanceName].InstanceKey = $windowsInstance.InstanceKey
                $candidates[$instanceName].ConfiguredNamedPipe = $windowsInstance.ConfiguredNamedPipe
                $candidates[$instanceName].TcpEnabled = $windowsInstance.TcpEnabled
                $candidates[$instanceName].NamedPipesEnabled = $windowsInstance.NamedPipesEnabled
                $candidates[$instanceName].ServiceState = $windowsInstance.ServiceState
                $candidates[$instanceName].ServiceStartMode = $windowsInstance.ServiceStartMode
                if (-not $candidates[$instanceName].Sources.Contains('Windows service/registry/listener')) {
                    $candidates[$instanceName].Sources.Add('Windows service/registry/listener')
                }
            }
        }
        catch {
            $diagnosticId = Write-Diagnostic -Level WARNING `
                -Context "Expanding SQL instances on '$hostName'" `
                -Message 'Windows service/registry discovery failed; CSV and SQL Browser results will still be used.' `
                -ErrorRecord $_
            Write-Warning "Windows instance discovery failed on '$hostName'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
        }
        finally {
            if ($session) {
                Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
            }
        }
    }

    foreach ($candidate in $candidates.Values) {
        $dataSources = [System.Collections.Generic.List[string]]::new()
        foreach ($port in $candidate.TcpPorts) {
            $directDataSource = New-ExpandedSqlDataSource -DiscoveryRow $candidate.DiscoveryRow `
                -InstanceName $candidate.InstanceName -TcpPort $port
            if ($directDataSource -and -not $dataSources.Contains($directDataSource)) {
                $dataSources.Add($directDataSource)
            }
        }

        foreach ($pipeName in $candidate.NamedPipes) {
            $namedPipeDataSource = New-RemoteNamedPipeDataSource `
                -DiscoveryRow $candidate.DiscoveryRow -PipeName $pipeName
            if ($namedPipeDataSource -and -not $dataSources.Contains($namedPipeDataSource)) {
                $dataSources.Add($namedPipeDataSource)
            }
        }

        $namedDataSource = New-ExpandedSqlDataSource -DiscoveryRow $candidate.DiscoveryRow `
            -InstanceName $candidate.InstanceName -TcpPort ''
        if ($namedDataSource -and -not $dataSources.Contains($namedDataSource)) {
            $dataSources.Add($namedDataSource)
        }
        if ($dataSources.Count -eq 0) { continue }

        $key = '{0}|{1}' -f $server.ServerKey, $candidate.InstanceName
        $instances[$key] = [pscustomobject]@{
            DiscoveryRow          = $candidate.DiscoveryRow
            DataSources           = $dataSources
            ServerKey             = $server.ServerKey
            InstanceName          = $candidate.InstanceName
            TcpPorts              = $candidate.TcpPorts
            NamedPipes            = $candidate.NamedPipes
            InstanceKey           = $candidate.InstanceKey
            ConfiguredNamedPipe   = $candidate.ConfiguredNamedPipe
            TcpEnabled            = $candidate.TcpEnabled
            NamedPipesEnabled     = $candidate.NamedPipesEnabled
            ServiceState          = $candidate.ServiceState
            ServiceStartMode      = $candidate.ServiceStartMode
            InstanceDiscoverySource = $candidate.Sources -join '; '
        }
    }
}

if ($instances.Count -eq 0) {
    throw 'No connectable SQL Server instances were found in the discovery CSVs, SQL Browser, or Windows service/registry discovery.'
}

$instanceCountByServer = @{}
foreach ($instance in $instances.Values) {
    if (-not $instanceCountByServer.ContainsKey($instance.ServerKey)) {
        $instanceCountByServer[$instance.ServerKey] = 0
    }
    $instanceCountByServer[$instance.ServerKey]++
}

$sqlAuthentication = if ($Credential) {
    "SQL credential ($($Credential.UserName))"
}
else {
    "Current Windows user ($([Security.Principal.WindowsIdentity]::GetCurrent().Name))"
}

$discoveryAuthentication = if ($InstanceDiscoveryAuthMode -eq 'BrowserOnly') {
    'BrowserOnly (no Windows authentication)'
}
elseif ($DiscoveryCredential) {
    "Windows credential ($($DiscoveryCredential.UserName))"
}
else {
    "Current Windows user ($([Security.Principal.WindowsIdentity]::GetCurrent().Name))"
}

Write-Host ''
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host ' SQL Server Database Storage Inventory' -ForegroundColor Cyan
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host (" Input files       : {0}" -f $loadedFiles.Count)
Write-Host (" Unique servers    : {0}" -f $servers.Count)
Write-Host (" Unique instances  : {0}" -f $instances.Count)
Write-Host (" SQL authentication: {0}" -f $sqlAuthentication)
Write-Host (" Instance discovery: {0}" -f $discoveryAuthentication)
Write-Host (" System databases  : {0}" -f $(if ($ExcludeSystemDatabases) { 'Excluded' } else { 'Included' }))
Write-Host (" Output CSV        : {0}" -f $OutputCsv)
Write-Host (" Diagnostic log    : {0}" -f $script:DiagnosticLogPath)
if ($skippedRows -gt 0) {
    $message = "Skipped $skippedRows discovery row(s) without a server name, FQDN, or IP address."
    $diagnosticId = Write-Diagnostic -Level WARNING -Context 'Discovery row validation' -Message $message
    Write-Warning "$message Diagnostic ID: $diagnosticId"
}
Write-Host ''

$results = [System.Collections.Generic.List[object]]::new()
$current = 0

foreach ($instance in ($instances.Values | Sort-Object ServerKey, InstanceName)) {
    $current++
    $row = $instance.DiscoveryRow
    $dataSource = $instance.DataSources[0]
    $sqlTcpPort = $instance.TcpPorts -join '; '
    $multipleSQLInstances = if ($instanceCountByServer[$instance.ServerKey] -gt 1) { 'Yes' } else { 'No' }
    $sqlVersion = Get-PropertyValue -InputObject $row -Name 'SQLVersion'
    $sqlRelease = Get-PropertyValue -InputObject $row -Name 'SQLRelease'
    $connection = $null
    $sqlVCoreCount = $null
    $databaseCount = $null
    $databasesMeasured = 0
    $dataUsedMb = 0.0
    $logUsedMb = 0.0
    $issues = [System.Collections.Generic.List[string]]::new()
    $status = 'OK'
    $connectivityDetail = ''
    $skipSqlConnection = $false

    if (-not [string]::IsNullOrWhiteSpace($instance.ServiceState) -and
        $instance.ServiceState -ne 'Running') {
        $status = 'STOPPED'
        $connectivityDetail = "SQL Server service state is $($instance.ServiceState)."
        $skipSqlConnection = $true
    }
    elseif ($instance.TcpEnabled -eq 0 -and $instance.NamedPipesEnabled -eq 0) {
        $status = 'NOT CONNECTABLE'
        $connectivityDetail = 'TCP/IP and Named Pipes are disabled; the instance accepts local connections only.'
        $diagnosticId = Write-Diagnostic -Level WARNING `
            -Context "Assessing connectivity for '$($instance.ServerKey)\$($instance.InstanceName)'" `
            -Message $connectivityDetail
        $issues.Add("[Diagnostic $diagnosticId] $connectivityDetail")
        Write-Warning "Skipping '$($instance.ServerKey)\$($instance.InstanceName)': $connectivityDetail Diagnostic ID: $diagnosticId"
        $skipSqlConnection = $true
    }

    Write-Progress -Activity 'Measuring SQL database storage' `
        -Status ("{0}\{1} ({2} of {3})" -f $instance.ServerKey, $instance.InstanceName,
            $current, $instances.Count) `
        -PercentComplete (($current / $instances.Count) * 100)

    if (-not $skipSqlConnection) {
    try {
        $lastConnectionError = $null
        foreach ($candidateDataSource in $instance.DataSources) {
            $dataSource = $candidateDataSource
            try {
                $connectionString = New-SqlConnectionString -DataSource $dataSource -Database 'master'
                $connection = [System.Data.SqlClient.SqlConnection]::new($connectionString)
                $connection.Open()
                break
            }
            catch {
                $lastConnectionError = $_
                if ($connection) {
                    $connection.Dispose()
                    $connection = $null
                }
            }
        }

        if (-not $connection) {
            $attemptedTargets = $instance.DataSources -join '; '
            $lastMessage = if ($lastConnectionError) {
                $lastConnectionError.Exception.Message
            }
            else {
                'No connection error was returned.'
            }
            $connectionException = [System.Exception]::new(
                "All connection targets failed ($attemptedTargets). Last error: $lastMessage",
                $(if ($lastConnectionError) { $lastConnectionError.Exception } else { $null })
            )
            throw $connectionException
        }

        if ($dataSource -match '(?i)^tcp:[^,]+,(?<Port>\d{1,5})$') {
            $sqlTcpPort = $Matches.Port
        }

        try {
            $serverInfo = Invoke-SqlDataTable -Connection $connection -Query @'
SELECT CONVERT(nvarchar(128), SERVERPROPERTY('ProductVersion')) AS SQLVersion;
'@
            $queriedVersion = [string] $serverInfo.Rows[0].SQLVersion
            if (-not [string]::IsNullOrWhiteSpace($queriedVersion)) {
                $sqlVersion = $queriedVersion
                $sqlRelease = Get-SqlRelease -Version $queriedVersion
            }
        }
        catch {
            $diagnosticId = Write-Diagnostic -Level ERROR `
                -Context "Reading SQL version on '$dataSource'" `
                -Message 'The SQL Server version query failed.' -ErrorRecord $_
            $issues.Add("[Diagnostic $diagnosticId] SQL version unavailable: $($_.Exception.Message)")
            Write-Warning "SQL version query failed on '$dataSource'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
        }

        try {
            $schedulerInfo = Invoke-SqlDataTable -Connection $connection -Query @'
SELECT COUNT_BIG(*) AS SQLVCoreCount
FROM sys.dm_os_schedulers
WHERE scheduler_id < 1048576
  AND status = 'VISIBLE ONLINE';
'@
            $sqlVCoreCount = [long] $schedulerInfo.Rows[0].SQLVCoreCount
        }
        catch {
            $diagnosticId = Write-Diagnostic -Level ERROR `
                -Context "Reading SQL-visible vCores on '$dataSource'" `
                -Message 'The SQL scheduler query failed.' -ErrorRecord $_
            $issues.Add("[Diagnostic $diagnosticId] SQL vCore count unavailable: $($_.Exception.Message)")
            Write-Warning "SQL vCore count failed on '$dataSource'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
        }

        $databaseFilter = if ($ExcludeSystemDatabases) {
            "WHERE name NOT IN ('master', 'model', 'msdb', 'tempdb')"
        }
        else {
            ''
        }

        $databaseQuery = @"
SELECT name, CONVERT(varchar(20), DATABASEPROPERTYEX(name, 'Status')) AS DatabaseStatus
FROM master.dbo.sysdatabases
$databaseFilter
ORDER BY name;
"@
        $databases = Invoke-SqlDataTable -Connection $connection -Query $databaseQuery
        $databaseCount = $databases.Rows.Count

        $logUsageByDatabase = @{}
        try {
            $connection.ChangeDatabase('master')
            $logUsage = Invoke-SqlDataTable -Connection $connection -Query 'DBCC SQLPERF(LOGSPACE) WITH NO_INFOMSGS;'
            foreach ($logRow in $logUsage.Rows) {
                $databaseName = [string] $logRow.'Database Name'
                $logSizeMb = [double] $logRow.'Log Size (MB)'
                $logPercentUsed = [double] $logRow.'Log Space Used (%)'
                $logUsageByDatabase[$databaseName] = $logSizeMb * ($logPercentUsed / 100.0)
            }
        }
        catch {
            $diagnosticId = Write-Diagnostic -Level ERROR `
                -Context "Reading transaction-log usage on '$dataSource'" `
                -Message 'DBCC SQLPERF(LOGSPACE) failed.' -ErrorRecord $_
            $issues.Add("[Diagnostic $diagnosticId] Transaction-log usage unavailable: $($_.Exception.Message)")
            Write-Warning "Transaction-log usage failed on '$dataSource'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
        }

        foreach ($database in $databases.Rows) {
            $databaseName = [string] $database.name
            $databaseStatus = [string] $database.DatabaseStatus

            if ($databaseStatus -ne 'ONLINE') {
                $message = "Database '$databaseName' is $databaseStatus"
                $diagnosticId = Write-Diagnostic -Level WARNING `
                    -Context "Measuring database on '$dataSource'" -Message $message
                $issues.Add("[Diagnostic $diagnosticId] $message")
                continue
            }

            try {
                $connection.ChangeDatabase($databaseName)
                $dataUsage = Invoke-SqlDataTable -Connection $connection -Query @'
SELECT COALESCE(
    SUM(CONVERT(float, FILEPROPERTY(name, 'SpaceUsed'))) * 8.0 / 1024.0,
    0.0
) AS UsedDataMB
FROM dbo.sysfiles
WHERE groupid > 0;
'@
                $dataUsedMb += [double] $dataUsage.Rows[0].UsedDataMB
                if ($logUsageByDatabase.ContainsKey($databaseName)) {
                    $logUsedMb += [double] $logUsageByDatabase[$databaseName]
                }
                else {
                    $issues.Add("Transaction-log usage unavailable for database '$databaseName'")
                }
                $databasesMeasured++
            }
            catch {
                $diagnosticId = Write-Diagnostic -Level ERROR `
                    -Context "Measuring database '$databaseName' on '$dataSource'" `
                    -Message 'The database usage query failed.' -ErrorRecord $_
                $issues.Add("[Diagnostic $diagnosticId] Database '$databaseName': $($_.Exception.Message)")
                Write-Warning "Database '$databaseName' failed on '$dataSource'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
            }
        }

        if ($issues.Count -gt 0) {
            $status = 'PARTIAL'
        }
    }
    catch {
        $status = 'FAILED'
        $diagnosticId = Write-Diagnostic -Level ERROR -Context "Connecting to or inventorying '$dataSource'" `
            -Message 'The SQL Server instance could not be inventoried.' -ErrorRecord $_
        $issues.Add("[Diagnostic $diagnosticId] $($_.Exception.Message)")
        $connectivityDetail = "Connection failed. Diagnostic ID: $diagnosticId"
        Write-Warning "Inventory failed for '$dataSource'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
    }
    finally {
        if ($connection) {
            try {
                $connection.Close()
                $connection.Dispose()
            }
            catch {
                $diagnosticId = Write-Diagnostic -Level ERROR -Context "Closing SQL connection to '$dataSource'" `
                    -Message 'The SQL connection could not be closed or disposed cleanly.' -ErrorRecord $_
                $issues.Add("[Diagnostic $diagnosticId] Connection cleanup failed: $($_.Exception.Message)")
                if ($status -eq 'OK') { $status = 'PARTIAL' }
                Write-Warning "Connection cleanup failed for '$dataSource'. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'."
            }
        }
    }
    }

    $totalUsedMb = $dataUsedMb + $logUsedMb
    $hasMeasurements = $status -in @('OK', 'PARTIAL')
    $results.Add([pscustomobject][ordered]@{
        ServerName          = Get-PropertyValue -InputObject $row -Name 'ServerName'
        InstanceName        = $instance.InstanceName
        MultipleSQLInstances = $multipleSQLInstances
        InstanceDiscoverySource = $instance.InstanceDiscoverySource
        FQDN                = Get-PropertyValue -InputObject $row -Name 'FQDN'
        IPAddress           = Get-PropertyValue -InputObject $row -Name 'IPAddress'
        DataSource          = $dataSource
        SQLInstanceKey      = $instance.InstanceKey
        SQLTcpPort          = $sqlTcpPort
        SQLNamedPipe        = $instance.ConfiguredNamedPipe
        SQLServiceState     = $instance.ServiceState
        SQLServiceStartMode = $instance.ServiceStartMode
        TCPEnabled          = if ($null -eq $instance.TcpEnabled) { 'Unknown' } elseif ($instance.TcpEnabled -eq 1) { 'Yes' } else { 'No' }
        NamedPipesEnabled   = if ($null -eq $instance.NamedPipesEnabled) { 'Unknown' } elseif ($instance.NamedPipesEnabled -eq 1) { 'Yes' } else { 'No' }
        ConnectivityDetail  = $connectivityDetail
        SQLRelease          = $sqlRelease
        SQLVersion          = $sqlVersion
        SQLVCoreCount       = $sqlVCoreCount
        DatabaseCount       = $databaseCount
        DatabasesMeasured   = $databasesMeasured
        DataUsedGB          = if ($hasMeasurements) { [Math]::Round($dataUsedMb / 1024.0, 3) } else { $null }
        LogUsedGB           = if ($hasMeasurements) { [Math]::Round($logUsedMb / 1024.0, 3) } else { $null }
        TotalStorageUsedGB  = if ($hasMeasurements) { [Math]::Round($totalUsedMb / 1024.0, 3) } else { $null }
        Status              = $status
        ScanTime            = Get-Date -Format 's'
    })
}

Write-Progress -Activity 'Measuring SQL database storage' -Completed

$sorted = @($results | Sort-Object ServerName, InstanceName)
$sorted | Format-Table ServerName, InstanceName, MultipleSQLInstances, SQLVCoreCount, DatabaseCount,
    DataUsedGB, TotalStorageUsedGB, Status -AutoSize

try {
    $outputDirectory = Split-Path -Parent $OutputCsv
    if ($outputDirectory -and -not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    $sorted | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host ''
    Write-Host ("CSV exported to : {0}" -f $OutputCsv) -ForegroundColor Green
}
catch {
    throw
}

$okCount = @($sorted | Where-Object { $_.Status -eq 'OK' }).Count
$partialCount = @($sorted | Where-Object { $_.Status -eq 'PARTIAL' }).Count
$failedCount = @($sorted | Where-Object { $_.Status -eq 'FAILED' }).Count
$stoppedCount = @($sorted | Where-Object { $_.Status -eq 'STOPPED' }).Count
$notConnectableCount = @($sorted | Where-Object { $_.Status -eq 'NOT CONNECTABLE' }).Count

Write-Host ("Successful        : {0}" -f $okCount)
Write-Host ("Partial           : {0}" -f $partialCount)
Write-Host ("Failed            : {0}" -f $failedCount)
Write-Host ("Stopped           : {0}" -f $stoppedCount)
Write-Host ("Not connectable   : {0}" -f $notConnectableCount)
Write-Host ("Elapsed           : {0:N1} seconds" -f ((Get-Date) - $script:StartTime).TotalSeconds)
Write-Host ''
}

try {
    Invoke-StorageInventory
}
catch {
    try {
        $diagnosticId = Write-Diagnostic -Level ERROR -Context 'Unhandled fatal script error' `
            -Message 'The inventory script terminated before successful completion.' -ErrorRecord $_
        Write-Error "Inventory terminated. Diagnostic ID: $diagnosticId. See '$script:DiagnosticLogPath'. $($_.Exception.Message)" `
            -ErrorAction Continue
    }
    catch {
        Write-Error "Inventory terminated and the diagnostic error could not be written: $($_.Exception.Message)" `
            -ErrorAction Continue
    }
    throw
}
```

---

## 10. Design notes

A few decisions worth explaining, in case you or a colleague later wonder *"why did he do it that way?"*

**Registry over SQL query.** Covered in section 2 — it removes the SQL-login dependency entirely, which was the single biggest cause of the original script returning nothing.

**Runspace pool over `ForEach-Object -Parallel`.** The `-Parallel` switch only exists in PowerShell 7. Most estates still run Windows PowerShell 5.1 on their servers, and a sequential /24 sweep at 700ms per host would take almost three minutes instead of thirteen seconds. The runspace pool gives identical performance on both.

**Never silently drop a host.** Every host that AD or the network scan identified produces a row, even when detail couldn't be read. A discovery report that quietly omits what it couldn't reach is worse than no report, because it looks complete.

**Version map is additive, not a filter.** An unknown major build is reported as `Unrecognised (major build N)` rather than skipped. This is the direct fix for the original `13.*` filter bug.

**Legacy paths are tried in order, not branched on a version guess.** Rather than detecting the SQL version first and then choosing a registry path, the scripts try each known layout in sequence and take the first that returns a build number. This means a layout that doesn't match any documented pattern still has a chance of resolving, and adding support for a new one is a single line in a list.

**`MaxHosts` safety brake.** A mistyped prefix (`/8` instead of `/28`) would otherwise try to probe 16 million addresses. The script refuses and tells you to raise the limit deliberately.

**Credential precedence.** Explicit `-Credential` beats `-AuthMode Prompt`, so the same script works interactively at a console and unattended in a scheduled task without editing anything.

**Separate discovery and SQL credentials in the inventory script.** WMI/registry instance expansion and the actual SQL connection are deliberately independent (`-DiscoveryCredential` vs. `-Credential`), because the account with admin rights on the Windows host is very often not the account that holds a SQL login — conflating them would have forced one identity to hold both, usually by over-granting.

**`NOT CONNECTABLE` over a misleading connection error.** When both TCP/IP and Named Pipes are disabled on a running instance, .NET's `SqlConnection` ultimately reports a generic error 26 ("Error Locating Server/Instance Specified"). The inventory script detects the disabled-protocol condition directly from the registry first and reports it as `NOT CONNECTABLE` with the real cause, rather than letting the generic SQL error stand as the explanation.

---

## 11. Files

| File | Description |
|---|---|
| `SQL-Discovery-DomainController.ps1` | AD SPN-based discovery script |
| `SQL-Discovery-MemberServer.ps1` | Subnet scan discovery script |
| `SQL-Database-Storage-Inventory.ps1` | Per-instance vCore, database, and storage inventory script |
| `SQL-Discovery-DomainController.sample.csv` | Sample output from the DC script |
| `SQL-Discovery-MemberServer.sample.csv` | Sample output from the subnet scan |
| `SQL-Database-Storage-Inventory.sample.csv` | Sample output from the storage inventory script |
| `SQL-Discovery-Toolkit.docx` | This document (Word) |
| `SQL-Discovery-Toolkit.md` | This document (Markdown source) |
| `SQL-Server-2016-End-of-Support-Discover-Decide-Modernise.pptx` | Supporting presentation |

---

*Russell McKee — https://www.linkedin.com/in/russellwbmckee/*
