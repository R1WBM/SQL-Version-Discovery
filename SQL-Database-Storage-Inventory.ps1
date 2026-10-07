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
