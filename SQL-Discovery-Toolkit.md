# SQL Server Discovery Toolkit

**Author:** Russell McKee
**LinkedIn:** https://www.linkedin.com/in/russellwbmckee/
**Version:** 2.1
**Date:** 25 September 2026

---

## Change log

| Version | Change |
|---|---|
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

This toolkit contains two complementary PowerShell scripts that inventory Microsoft SQL Server instances across an estate and export the results to CSV.

| Script | Runs from | Finds SQL by | Best for |
|---|---|---|---|
| `SQL-Discovery-DomainController.ps1` | A domain controller (or any host with RSAT) | Querying Active Directory for `MSSQLSvc` SPNs | Fast, authoritative sweep of everything registered in AD |
| `SQL-Discovery-MemberServer.ps1` | Any domain member server | Sweeping every IP in a subnet | Finding instances AD doesn't know about |

Both capture the same core fields: **Server Name, Instance, SQL Version, SQL Edition, FQDN and IP Address** — for **any SQL Server release from 6.5 / 7.0 / 2000 through to 2025**, with no hard-coded version filter.

### Why two scripts?

They answer different questions, and the gap between them is the interesting part.

The **domain controller** script asks *"what does Active Directory think we have?"*. It's near-instant (1.3 seconds in the lab) because it's a single LDAP query, but it only sees instances that registered a Service Principal Name. An instance running under `LocalSystem`, or one whose SPN registration failed, is invisible to it.

The **member server** script asks *"what's actually listening on the wire?"*. It's slower (13 seconds for a /24) but it doesn't care about AD registration — if something answers on a SQL port, it gets found.

**Run both and compare.** Anything the subnet scan finds that the AD scan missed is either an unregistered instance or shadow IT, and both are worth knowing about.

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

---

## 9. Full script listing

### 9.1 `SQL-Discovery-DomainController.ps1`

```powershell
<#
================================================================================
 SQL Server Discovery - Domain Controller Edition
================================================================================

 Author   : Russell McKee
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

---

## 11. Files

| File | Description |
|---|---|
| `SQL-Discovery-DomainController.ps1` | AD SPN-based discovery script |
| `SQL-Discovery-MemberServer.ps1` | Subnet scan discovery script |
| `SQL-Discovery-DomainController.sample.csv` | Sample output from the DC script |
| `SQL-Discovery-MemberServer.sample.csv` | Sample output from the subnet scan |
| `SQL-Discovery-Toolkit.docx` | This document (Word) |
| `SQL-Discovery-Toolkit.md` | This document (Markdown source) |
| `SQL-Server-2016-End-of-Support-Discover-Decide-Modernise.pptx` | Supporting presentation |

---

*Russell McKee — https://www.linkedin.com/in/russellwbmckee/*
