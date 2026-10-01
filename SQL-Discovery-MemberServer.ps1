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
