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
